const REPORT_PATH = "../reports/stock_closing_prices.csv";
const PNL_REPORT_PATH = "../reports/stock_pnl_summary.csv";
const MTF_PNL_REPORT_PATH = "../reports/mtf_pnl_summary.csv";
const LEDGER_PATH = "../reports/exported_all_ledger.csv";
const DAY_COUNT = 30;
const IS_FILE_PROTOCOL = window.location.protocol === "file:";
const VIEW = {
  CHARTS: "charts",
  PNL: "pnl",
  MTF: "mtf",
};
const STOCK_FILTER = {
  ALL: "all",
  REALIZED: "realized",
  UNREALIZED: "unrealized",
};
const PNL_BASIS = {
  ALL_HOLDINGS: "all_holdings",
  PERSONAL_FUNDING: "personal_funding",
};
const TABLE_KEYS = {
  PNL: "pnl",
  MTF: "mtf",
};

const searchInput = document.getElementById("searchInput");
const statusBanner = document.getElementById("statusBanner");
const cardsContainer = document.getElementById("cardsContainer");
const template = document.getElementById("stockCardTemplate");

const stocksLoaded = document.getElementById("stocksLoaded");
const stocksInvested = document.getElementById("stocksInvested");
const pnlFilterSwitch = document.getElementById("pnlFilterSwitch");
const pnlFilterButtons = Array.from(pnlFilterSwitch.querySelectorAll(".pnl-filter-button"));
const pnlBasisSwitch = document.getElementById("pnlBasisSwitch");
const pnlBasisButtons = Array.from(pnlBasisSwitch.querySelectorAll(".pnl-basis-button"));
const mtfFilterSwitch = document.getElementById("mtfFilterSwitch");
const mtfFilterButtons = Array.from(mtfFilterSwitch.querySelectorAll(".mtf-filter-button"));
const chartViewButton = document.getElementById("chartViewButton");
const pnlViewButton = document.getElementById("pnlViewButton");
const mtfViewButton = document.getElementById("mtfViewButton");
const chartsPage = document.getElementById("chartsPage");
const pnlPage = document.getElementById("pnlPage");
const mtfPage = document.getElementById("mtfPage");
const pnlTable = document.getElementById("pnlTable");
const mtfTable = document.getElementById("mtfTable");
const pnlTableBody = document.getElementById("pnlTableBody");
const mtfTableBody = document.getElementById("mtfTableBody");
const pnlExcludedColumns = document.getElementById("pnlExcludedColumns");
const mtfExcludedColumns = document.getElementById("mtfExcludedColumns");
const totalRealizedPnl = document.getElementById("totalRealizedPnl");
const totalChargesTaxesOthers = document.getElementById("totalChargesTaxesOthers");
const netRealizedPnl = document.getElementById("netRealizedPnl");
const totalUnrealizedPnl = document.getElementById("totalUnrealizedPnl");
const totalOtherCreditsDebits = document.getElementById("totalOtherCreditsDebits");
const netTotalPnl = document.getElementById("netTotalPnl");
const totalAmountInvested = document.getElementById("totalAmountInvested");
const mtfHoldingPnl = document.getElementById("mtfHoldingPnl");
const mtfNetFundingPnl = document.getElementById("mtfNetFundingPnl");
const mtfYourFunding = document.getElementById("mtfYourFunding");
const mtfTotalCarryingCost = document.getElementById("mtfTotalCarryingCost");
const timelineTotalAdded = document.getElementById("timelineTotalAdded");
const timelineTotalWithdrawn = document.getElementById("timelineTotalWithdrawn");
const timelineCurrentGain = document.getElementById("timelineCurrentGain");

let allStocks = [];
let pnlRows = [];
let mtfRows = [];
let ledgerRows = [];
let otherCreditsDebitsValue = 0;
let portfolioChart = null;
let activeView = VIEW.CHARTS;
let activePnlFilter = STOCK_FILTER.ALL;
let activePnlBasis = PNL_BASIS.ALL_HOLDINGS;
let activeMtfFilter = STOCK_FILTER.ALL;
const hiddenColumnsByTable = {
  [TABLE_KEYS.PNL]: new Set(),
  [TABLE_KEYS.MTF]: new Set(),
};

searchInput.addEventListener("input", () => {
  if (activeView === VIEW.CHARTS) {
    renderCards(searchInput.value);
  } else if (activeView === VIEW.PNL) {
    refreshPnlView();
  } else {
    refreshMtfView();
  }
});

chartViewButton.addEventListener("click", () => {
  setView(VIEW.CHARTS);
});

pnlViewButton.addEventListener("click", () => {
  setView(VIEW.PNL);
});

mtfViewButton.addEventListener("click", () => {
  setView(VIEW.MTF);
});

pnlFilterButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setPnlFilter(button.dataset.filter || STOCK_FILTER.ALL);
  });
});

pnlBasisButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setPnlBasis(button.dataset.basis || PNL_BASIS.ALL_HOLDINGS);
  });
});

mtfFilterButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setMtfFilter(button.dataset.filter || STOCK_FILTER.ALL);
  });
});

loadDashboard();
initializeColumnControls();

async function loadDashboard() {
  setStatus("", false);

  try {
    const { reportCsvText, pnlCsvText, mtfCsvText, ledgerCsvText } = await loadCsvSources();
    const rows = parseCsv(reportCsvText);
    const recommendationDateBySymbol = buildRecommendationDateBySymbol(rows);
    const mtfSourceRows = mtfCsvText ? parseCsv(mtfCsvText) : [];
    const mtfFundingByKey = buildMtfFundingByKey(mtfSourceRows);
    allStocks = rows.map(transformRow).filter((item) => item.points.length > 0);
    pnlRows = parseCsv(pnlCsvText).map((row) =>
      transformPnlRow(row, recommendationDateBySymbol, mtfFundingByKey)
    );
    if (mtfCsvText) {
      mtfRows = mtfSourceRows.map((row) => transformMtfRow(row));
    } else {
      mtfRows = [];
    }
    ledgerRows = ledgerCsvText ? parseCsv(ledgerCsvText) : [];
    otherCreditsDebitsValue = computeOtherCreditsDebits(ledgerRows);

    updateSummary(allStocks);
    renderCards(searchInput.value);
    refreshPnlView();
    refreshMtfView();
    renderInvestmentTimeline(getCurrentPnlRows());

    if (allStocks.length === 0) {
      setStatus("The report has no chartable rows yet. Run python/fetch_yahoo_prices.py to refresh CSV reports.", false);
    }
  } catch (error) {
    allStocks = [];
    pnlRows = [];
    mtfRows = [];
    ledgerRows = [];
    otherCreditsDebitsValue = 0;
    updateSummary([]);
    updatePnlSummary([]);
    updateMtfSummary([]);
    cardsContainer.innerHTML = "";
    pnlTableBody.innerHTML = "";
    mtfTableBody.innerHTML = "";
    renderInvestmentTimeline([]);
    setStatus(
      `Unable to load report files. ${error.message || ""} Ensure reports are generated and committed, then open the dashboard URL.`,
      true
    );
    console.error(error);
  }
}

function getCurrentPnlRows() {
  return getPnlRowsByActiveFilter(pnlRows);
}

function getCurrentMtfRows() {
  return getMtfRowsByActiveFilter(mtfRows);
}

function refreshPnlView() {
  updatePnlSummary(getCurrentPnlRows());
  renderPnlTable(searchInput.value);
}

function refreshMtfView() {
  updateMtfSummary(getCurrentMtfRows());
  renderMtfTable(searchInput.value);
}

async function loadCsvSources() {
  if (IS_FILE_PROTOCOL) {
    const embeddedSources = getEmbeddedCsvSources();
    if (embeddedSources) {
      return embeddedSources;
    }
    throw new Error(
      "Local file mode needs dashboard/data.js. Run ./shell/update_data.sh once so data.js is generated."
    );
  }

  try {
    return await fetchCsvSources();
  } catch (fetchError) {
    const embeddedSources = getEmbeddedCsvSources();
    if (embeddedSources) {
      return embeddedSources;
    }
    throw fetchError;
  }
}

function getEmbeddedCsvSources() {
  const embeddedData = readEmbeddedData();
  if (!embeddedData.stockClosingPricesCsv || !embeddedData.stockPnlSummaryCsv) {
    return null;
  }
  return {
    reportCsvText: embeddedData.stockClosingPricesCsv,
    pnlCsvText: embeddedData.stockPnlSummaryCsv,
    mtfCsvText: embeddedData.mtfPnlSummaryCsv,
    ledgerCsvText: embeddedData.allLedgerCsv,
  };
}

async function fetchCsvSources() {
  const cacheBuster = Date.now();
  const [reportResponse, pnlResponse, mtfResponse, ledgerResponse] = await Promise.all([
    fetch(`${REPORT_PATH}?t=${cacheBuster}`),
    fetch(`${PNL_REPORT_PATH}?t=${cacheBuster}`),
    fetch(`${MTF_PNL_REPORT_PATH}?t=${cacheBuster}`),
    fetch(`${LEDGER_PATH}?t=${cacheBuster}`),
  ]);
  if (!reportResponse.ok || !pnlResponse.ok) {
    const failures = [];
    if (!reportResponse.ok) {
      failures.push(`${REPORT_PATH} [${reportResponse.status}]`);
    }
    if (!pnlResponse.ok) {
      failures.push(`${PNL_REPORT_PATH} [${pnlResponse.status}]`);
    }
    throw new Error(`Could not load: ${failures.join(", ")}`);
  }

  const [reportCsvText, pnlCsvText] = await Promise.all([reportResponse.text(), pnlResponse.text()]);
  const mtfCsvText = mtfResponse.ok ? await mtfResponse.text() : "";
  const ledgerCsvText = ledgerResponse.ok ? await ledgerResponse.text() : "";
  return { reportCsvText, pnlCsvText, mtfCsvText, ledgerCsvText };
}

function readEmbeddedData() {
  const data = window.__DASHBOARD_DATA__;
  if (!data || typeof data !== "object") {
    return {
      stockClosingPricesCsv: "",
      stockPnlSummaryCsv: "",
      mtfPnlSummaryCsv: "",
      allLedgerCsv: "",
    };
  }

  return {
    stockClosingPricesCsv: String(data.stockClosingPricesCsv || ""),
    stockPnlSummaryCsv: String(data.stockPnlSummaryCsv || ""),
    mtfPnlSummaryCsv: String(data.mtfPnlSummaryCsv || ""),
    allLedgerCsv: String(data.allLedgerCsv || ""),
  };
}

function updateSummary(items) {
  stocksLoaded.textContent = String(items.length);
  const investedCount = pnlRows.filter(
    (item) => item.buyQuantityValue > 0 && item.sellQuantityValue === 0
  ).length;
  stocksInvested.textContent = String(investedCount);
}

function renderCards(filterText = "") {
  const normalizedFilter = filterText.trim().toLowerCase();
  const filtered = allStocks
    .filter((item) => item.stockCode.toLowerCase().includes(normalizedFilter))
    .sort((left, right) => {
      const leftTime = Date.parse(left.recommendationDate || "") || 0;
      const rightTime = Date.parse(right.recommendationDate || "") || 0;
      const dateCompare = rightTime - leftTime;
      if (dateCompare !== 0) {
        return dateCompare;
      }
      return left.stockCode.localeCompare(right.stockCode);
    });

  cardsContainer.innerHTML = "";

  if (filtered.length === 0) {
    const emptyState = document.createElement("div");
    emptyState.className = "empty-state";
    emptyState.textContent = normalizedFilter
      ? "No stocks match this filter."
      : "No stock data available to chart yet.";
    cardsContainer.appendChild(emptyState);
    return;
  }

  filtered.forEach((item, index) => {
    const fragment = template.content.cloneNode(true);
    const card = fragment.querySelector(".stock-card");
    const investment = findInvestment(item.stockCode);
    const positionStatus = findPositionStatus(item.stockCode);
    const stockBadges = fragment.querySelector(".stock-badges");

    fragment.querySelector(".stock-code").textContent = item.stockCode;
    fragment.querySelector(".stock-date").textContent = `Recommendation Date: ${item.recommendationDate}`;

    const investmentSummary = fragment.querySelector(".investment-summary");
    stockBadges.remove();

    if (investment) {
      investmentSummary.classList.add("visible");
      fragment.querySelector(".invested-amount").textContent = investment.buyValue;
      const investedReturn = fragment.querySelector(".invested-return");
      investedReturn.textContent = investment.returnPct || "--";
      const investedReturnClass =
        investment.returnPctValue > 0 ? "table-positive" : investment.returnPctValue < 0 ? "table-negative" : null;
      if (investedReturnClass) {
        investedReturn.classList.add(investedReturnClass);
      }
    } else {
      investmentSummary.remove();
    }

    const latestPoint = item.points[item.points.length - 1];
    fragment.querySelector(".latest-price").textContent = latestPoint.price;
    fragment.querySelector(".position-status").textContent = positionStatus;
    fragment.querySelector(".highest-price").textContent = item.highestPrice || "--";
    fragment.querySelector(".target-1-hit").innerHTML = renderTargetStatus(
      item.hitTarget1,
      item.target1,
      item.target1ReturnPct
    );
    fragment.querySelector(".target-2-hit").innerHTML = renderTargetStatus(
      item.hitTarget2,
      item.target2,
      item.target2ReturnPct
    );
    fragment.querySelector(".target-3-hit").innerHTML = renderTargetStatus(
      item.hitTarget3,
      item.target3,
      item.target3ReturnPct
    );
    fragment.querySelector(".target-4-hit").innerHTML = renderTargetStatus(
      item.hitTarget4,
      item.target4,
      item.target4ReturnPct
    );

    drawChart(fragment.querySelector(".chart"), item.points);
    fillTable(fragment.querySelector(".data-table-body"), item.points);

    card.style.animationDelay = `${index * 70}ms`;
    cardsContainer.appendChild(fragment);
  });
}

function renderPnlTable(filterText = "") {
  const normalizedFilter = filterText.trim().toLowerCase();
  const filtered = getPnlRowsByActiveFilter(pnlRows)
    .filter((item) => item.stockCode.toLowerCase().includes(normalizedFilter))
    .sort((left, right) => {
      const leftDate = left.tradeDate || "9999-12-31";
      const rightDate = right.tradeDate || "9999-12-31";
      const dateCompare = rightDate.localeCompare(leftDate);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      return left.stockCode.localeCompare(right.stockCode);
    });
  pnlTableBody.innerHTML = "";

  if (filtered.length === 0) {
    const row = document.createElement("tr");
    row.innerHTML = `<td colspan="${getVisibleColumnCount(TABLE_KEYS.PNL)}">${normalizedFilter ? "No P&L rows match this filter." : "No P&L data available yet."}</td>`;
    pnlTableBody.appendChild(row);
    applyColumnVisibility(TABLE_KEYS.PNL);
    return;
  }

  filtered.forEach((item) => {
    const returnBaseValue = getPnlRowBaseValue(item);
    const returnPctValue = returnBaseValue ? (item.totalPnlValue / returnBaseValue) * 100 : null;
    const returnPctText = returnPctValue === null ? "--" : `${returnPctValue.toFixed(2)}%`;
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${item.stockCode}</td>
      <td>${item.tradeDate || "--"}</td>
      <td>${item.buyQuantity}</td>
      <td>${item.averageBuyPrice || "--"}</td>
      <td>${item.buyValue || "--"}</td>
      <td>${item.sellQuantity}</td>
      <td>${item.averageSellPrice || "--"}</td>
      <td>${item.sellValue || "--"}</td>
      <td>${item.netQuantity}</td>
      <td>${item.latestMarketPrice || "--"}</td>
      <td class="${item.realizedPnlValue >= 0 ? "table-positive" : "table-negative"}">${item.realizedPnl}</td>
      <td class="${item.chargesTaxesOthersValue > 0 ? "table-negative" : ""}">${item.chargesTaxesOthers}</td>
      <td class="${item.netRealizedPnlValue >= 0 ? "table-positive" : "table-negative"}">${item.netRealizedPnl}</td>
      <td class="${item.unrealizedPnlValue >= 0 ? "table-positive" : "table-negative"}">${item.unrealizedPnl}</td>
      <td class="${item.totalPnlValue >= 0 ? "table-positive" : "table-negative"}">${item.totalPnl}</td>
      <td class="${returnPctValue === null ? "" : returnPctValue >= 0 ? "table-positive" : "table-negative"}">${returnPctText}</td>
    `;
    pnlTableBody.appendChild(row);
  });
  applyColumnVisibility(TABLE_KEYS.PNL);
}

function renderMtfTable(filterText = "") {
  const normalizedFilter = filterText.trim().toLowerCase();
  const filtered = getMtfRowsByActiveFilter(mtfRows)
    .filter((item) => item.stockCode.toLowerCase().includes(normalizedFilter))
    .sort((left, right) => {
      const leftDate = left.tradeDate || "9999-12-31";
      const rightDate = right.tradeDate || "9999-12-31";
      const dateCompare = rightDate.localeCompare(leftDate);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      return left.stockCode.localeCompare(right.stockCode);
    });
  mtfTableBody.innerHTML = "";

  if (filtered.length === 0) {
    const row = document.createElement("tr");
    row.innerHTML = `<td colspan="${getVisibleColumnCount(TABLE_KEYS.MTF)}">${normalizedFilter ? "No MTF rows match this filter." : "No MTF P&L data available yet."}</td>`;
    mtfTableBody.appendChild(row);
    applyColumnVisibility(TABLE_KEYS.MTF);
    return;
  }

  filtered.forEach((item) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${item.stockCode}</td>
      <td>${item.tradeDate || "--"}</td>
      <td>${item.buyQuantity || "--"}</td>
      <td>${item.buyValue || "--"}</td>
      <td>${item.yourFunding || "--"}</td>
      <td>${item.zerodhaFunding || "--"}</td>
      <td>${item.leverage || "--"}</td>
      <td>${item.sellQuantity || "--"}</td>
      <td>${item.sellValue || "--"}</td>
      <td>${item.netQuantity || "--"}</td>
      <td>${item.latestMarketPrice || "--"}</td>
      <td class="${item.holdingPnlValue >= 0 ? "table-positive" : "table-negative"}">${item.holdingPnl}</td>
      <td class="${item.holdingReturnPct ? (item.holdingReturnPctValue >= 0 ? "table-positive" : "table-negative") : ""}">${item.holdingReturnPct || "--"}</td>
      <td class="${item.fundingReturnPct ? (item.fundingReturnPctValue >= 0 ? "table-positive" : "table-negative") : ""}">${item.fundingReturnPct || "--"}</td>
      <td class="${item.mtfInterestCostValue > 0 ? "table-negative" : ""}">${item.mtfInterestCost}</td>
      <td class="${item.mtfPledgeChargesValue > 0 ? "table-negative" : ""}">${item.mtfPledgeCharges}</td>
      <td class="${item.totalCarryingCostValue > 0 ? "table-negative" : ""}">${item.totalCarryingCost}</td>
      <td class="${item.netFundingPnlValue >= 0 ? "table-positive" : "table-negative"}">${item.netFundingPnl}</td>
      <td class="${item.netFundingReturnPct ? (item.netFundingReturnPctValue >= 0 ? "table-positive" : "table-negative") : ""}">${item.netFundingReturnPct || "--"}</td>
    `;
    mtfTableBody.appendChild(row);
  });
  applyColumnVisibility(TABLE_KEYS.MTF);
}

function fillTable(tableBody, points) {
  tableBody.innerHTML = "";
  points.forEach((point) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>Day ${point.dayIndex}</td>
      <td>${point.date}</td>
      <td>${point.price}</td>
      <td class="${point.returnValue >= 0 ? "table-positive" : "table-negative"}">${point.returnText}</td>
    `;
    tableBody.appendChild(row);
  });
}

function drawChart(svg, points) {
  const width = 640;
  const height = 260;
  const padding = { top: 18, right: 16, bottom: 34, left: 54 };
  const chartWidth = width - padding.left - padding.right;
  const chartHeight = height - padding.top - padding.bottom;

  const minReturn = Math.min(...points.map((point) => point.returnValue), 0);
  const maxReturn = Math.max(...points.map((point) => point.returnValue), 0);
  const span = maxReturn - minReturn || 1;
  const yTicks = 5;

  const xForIndex = (index) =>
    padding.left + (points.length === 1 ? chartWidth / 2 : (index / (points.length - 1)) * chartWidth);
  const yForValue = (value) => padding.top + ((maxReturn - value) / span) * chartHeight;

  const linePath = points
    .map((point, index) => `${index === 0 ? "M" : "L"} ${xForIndex(index).toFixed(2)} ${yForValue(point.returnValue).toFixed(2)}`)
    .join(" ");

  const areaPath = `
    ${linePath}
    L ${xForIndex(points.length - 1).toFixed(2)} ${(padding.top + chartHeight).toFixed(2)}
    L ${xForIndex(0).toFixed(2)} ${(padding.top + chartHeight).toFixed(2)}
    Z
  `;

  const gridLines = Array.from({ length: yTicks }, (_, tickIndex) => {
    const ratio = tickIndex / (yTicks - 1);
    const value = maxReturn - ratio * span;
    const y = yForValue(value);
    return `
      <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="rgba(255,255,255,0.07)" stroke-dasharray="4 6" />
      <text x="${padding.left - 10}" y="${y + 4}" text-anchor="end" font-size="11" fill="#5a6b8a">${formatPercent(value)}</text>
    `;
  }).join("");

  const zeroLineY = yForValue(0);
  const pointsMarkup = points
    .map((point, index) => {
      const x = xForIndex(index);
      const y = yForValue(point.returnValue);
      return `<circle cx="${x}" cy="${y}" r="4.5" fill="${point.returnValue >= 0 ? "#00e5a0" : "#ff4d6a"}" />`;
    })
    .join("");

  const xLabels = points
    .filter((_, index) => index === 0 || index === points.length - 1 || index === Math.floor((points.length - 1) / 2))
    .map((point) => {
      const x = xForIndex(point.dayIndex - 1);
      return `<text x="${x}" y="${height - 10}" text-anchor="middle" font-size="11" fill="#5a6b8a">${point.date.slice(5)}</text>`;
    })
    .join("");

  svg.innerHTML = `
    <defs>
      <linearGradient id="areaGradient" x1="0" x2="0" y1="0" y2="1">
        <stop offset="0%" stop-color="rgba(0,229,160,0.22)" />
        <stop offset="100%" stop-color="rgba(0,229,160,0.03)" />
      </linearGradient>
    </defs>
    <rect x="0" y="0" width="${width}" height="${height}" rx="18" fill="transparent" />
    ${gridLines}
    <line x1="${padding.left}" y1="${zeroLineY}" x2="${width - padding.right}" y2="${zeroLineY}" stroke="rgba(255,255,255,0.15)" />
    <path d="${areaPath}" fill="url(#areaGradient)" />
    <path d="${linePath}" fill="none" stroke="#00e5a0" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" />
    ${pointsMarkup}
    ${xLabels}
  `;
}

function transformRow(row) {
  const stockCode = row["Stock Code/Name"] || "";
  const recommendationDate = row["Recommendation Date"] || "";
  const points = [];
  const inputTarget1 = parseOptionalNumber(row["Target 1"]);
  const inputTarget2 = parseOptionalNumber(row["Target 2"]);
  const inputTarget3 = parseOptionalNumber(row["Target 3"]);
  const inputTarget4 = parseOptionalNumber(row["Target 4"]);
  const hasFourTargets = inputTarget3 !== null && inputTarget4 !== null;
  const splitTargets = hasFourTargets
    ? [inputTarget1, inputTarget2, inputTarget3, inputTarget4]
    : distributeTargets(inputTarget1, inputTarget2);
  const buyPriceRecommendationValue = parseNumber(row["Buy Price Recommendation"]);
  const highestPriceValue = parseNumber(row["Highest Price (30D)"]);

  for (let index = 1; index <= DAY_COUNT; index += 1) {
    const date = row[`Date ${index}`];
    const price = row[`Day ${index} Price`];
    const returnText = row[`Day ${index} Return %`];
    if (!date || !price || !returnText) {
      continue;
    }

    points.push({
      dayIndex: index,
      date,
      price,
      returnText,
      returnValue: parsePercent(returnText),
    });
  }

  return {
    stockCode,
    normalizedSymbol: normalizeSymbol(stockCode),
    recommendationDate,
    target1: formatTarget(splitTargets[0]),
    target2: formatTarget(splitTargets[1]),
    target3: formatTarget(splitTargets[2]),
    target4: formatTarget(splitTargets[3]),
    buyPriceRecommendation: row["Buy Price Recommendation"] || "",
    highestPrice: row["Highest Price (30D)"] || "",
    hitTarget1: classifyTargetHit(highestPriceValue, splitTargets[0]),
    hitTarget2: classifyTargetHit(highestPriceValue, splitTargets[1]),
    hitTarget3: classifyTargetHit(highestPriceValue, splitTargets[2]),
    hitTarget4: classifyTargetHit(highestPriceValue, splitTargets[3]),
    target1ReturnPct: formatTargetReturnPct(splitTargets[0], buyPriceRecommendationValue),
    target2ReturnPct: formatTargetReturnPct(splitTargets[1], buyPriceRecommendationValue),
    target3ReturnPct: formatTargetReturnPct(splitTargets[2], buyPriceRecommendationValue),
    target4ReturnPct: formatTargetReturnPct(splitTargets[3], buyPriceRecommendationValue),
    points,
  };
}

function transformPnlRow(row, recommendationDateBySymbol = {}, mtfFundingByKey = {}) {
  const normalizedStockCode = normalizeSymbol(row["Stock Code/Name"] || "");
  const normalizedMatchedSymbol = normalizeSymbol(row["Matched Report Symbol"] || "");
  const tradeDate = row["Trade Date"] || "";
  const fundingKey =
    buildSymbolTradeKey(normalizedMatchedSymbol, tradeDate) ||
    buildSymbolTradeKey(normalizedStockCode, tradeDate);
  const mtfFundingValue = fundingKey ? mtfFundingByKey[fundingKey] : undefined;
  const hasMtfFunding = Number.isFinite(mtfFundingValue);
  const buyValueRaw = parseNumber(row["Buy Value"]);
  const recommendationDate =
    recommendationDateBySymbol[normalizedMatchedSymbol] ||
    recommendationDateBySymbol[normalizedStockCode] ||
    "";

  return {
    stockCode: row["Stock Code/Name"] || "",
    tradeDate,
    matchedReportSymbol: row["Matched Report Symbol"] || "",
    recommendationDate,
    buyQuantity: row["Buy Quantity"] || "",
    averageBuyPrice: row["Average Buy Price"] || "",
    buyQuantityValue: parseNumber(row["Buy Quantity"]),
    buyValue: formatCurrency(buyValueRaw),
    buyValueRaw,
    myFundingRaw: hasMtfFunding ? mtfFundingValue : buyValueRaw,
    hasMtfFunding,
    sellQuantity: row["Sell Quantity"] || "",
    sellQuantityValue: parseNumber(row["Sell Quantity"]),
    averageSellPrice: row["Average Sell Price"] || "",
    sellValue: formatCurrency(parseNumber(row["Sell Value"])),
    sellValueRaw: parseNumber(row["Sell Value"]),
    netQuantity: row["Net Quantity"] || "",
    netQuantityValue: parseNumber(row["Net Quantity"]),
    latestMarketPrice: row["Latest Market Price"] || "",
    latestMarketDate: row["Latest Market Date"] || "",
    realizedPnl: formatCurrency(parseNumber(row["Realized P&L"])),
    chargesTaxesOthers: formatCurrency(parseNumber(row["Charges, Taxes, Others"])),
    netRealizedPnl: formatCurrency(parseNumber(row["Net Realized P&L"])),
    unrealizedPnl: formatCurrency(parseNumber(row["Unrealized P&L"])),
    totalPnl: formatCurrency(parseNumber(row["Total P&L"])),
    returnPct: row["Return %"] || "",
    returnPctValue: parsePercent(row["Return %"] || ""),
    realizedPnlValue: parseNumber(row["Realized P&L"]),
    chargesTaxesOthersValue: parseNumber(row["Charges, Taxes, Others"]),
    netRealizedPnlValue: parseNumber(row["Net Realized P&L"]),
    unrealizedPnlValue: parseNumber(row["Unrealized P&L"]),
    totalPnlValue: parseNumber(row["Total P&L"]),
  };
}

function buildMtfFundingByKey(rows) {
  const fundingByKey = {};
  rows.forEach((row) => {
    const tradeDate = row["Trade Date"] || "";
    const funding = parseNumber(row["Your Funding"]);
    if (!tradeDate || !funding) {
      return;
    }
    const primaryKey = buildSymbolTradeKey(row["Stock Code/Name"], tradeDate);
    const secondaryKey = buildSymbolTradeKey(row["Matched Report Symbol"], tradeDate);
    const uniqueKeys = Array.from(new Set([primaryKey, secondaryKey].filter(Boolean)));
    uniqueKeys.forEach((key) => {
      fundingByKey[key] = (fundingByKey[key] || 0) + funding;
    });
  });
  return fundingByKey;
}

function buildSymbolTradeKey(symbol, tradeDate) {
  const normalizedSymbol = normalizeSymbol(symbol || "");
  const normalizedTradeDate = String(tradeDate || "").trim();
  if (!normalizedSymbol || !normalizedTradeDate) {
    return "";
  }
  return `${normalizedSymbol}|${normalizedTradeDate}`;
}

function transformMtfRow(row) {
  return {
    stockCode: row["Stock Code/Name"] || "",
    tradeDate: row["Trade Date"] || "",
    matchedReportSymbol: row["Matched Report Symbol"] || "",
    buyQuantity: row["Buy Quantity"] || "",
    buyValue: formatCurrency(parseNumber(row["Buy Value"])),
    buyValueRaw: parseNumber(row["Buy Value"]),
    yourFunding: formatCurrency(parseNumber(row["Your Funding"])),
    yourFundingRaw: parseNumber(row["Your Funding"]),
    zerodhaFunding: formatCurrency(parseNumber(row["Zerodha Funding"])),
    zerodhaFundingRaw: parseNumber(row["Zerodha Funding"]),
    leverage: row["Leverage X"] || "",
    sellQuantity: row["Sell Quantity"] || "",
    sellQuantityValue: parseNumber(row["Sell Quantity"]),
    sellValue: formatCurrency(parseNumber(row["Sell Value"])),
    sellValueRaw: parseNumber(row["Sell Value"]),
    netQuantity: row["Net Quantity"] || "",
    netQuantityValue: parseNumber(row["Net Quantity"]),
    latestMarketPrice: row["Latest Market Price"] || "",
    latestMarketDate: row["Latest Market Date"] || "",
    holdingPnl: formatCurrency(parseNumber(row["Holding P&L"])),
    holdingPnlValue: parseNumber(row["Holding P&L"]),
    holdingReturnPct: row["Holding Return %"] || "",
    holdingReturnPctValue: parsePercent(row["Holding Return %"] || ""),
    fundingReturnPct: row["Funding Return %"] || "",
    fundingReturnPctValue: parsePercent(row["Funding Return %"] || ""),
    mtfInterestCost: formatCurrency(-Math.abs(parseNumber(row["MTF Interest Cost"]))),
    mtfInterestCostValue: parseNumber(row["MTF Interest Cost"]),
    mtfPledgeCharges: formatCurrency(-Math.abs(parseNumber(row["MTF Pledge Charges"]))),
    mtfPledgeChargesValue: parseNumber(row["MTF Pledge Charges"]),
    totalCarryingCost: formatCurrency(-Math.abs(parseNumber(row["Total Carrying Cost"]))),
    totalCarryingCostValue: parseNumber(row["Total Carrying Cost"]),
    netFundingPnl: formatCurrency(parseNumber(row["Net Funding P&L"])),
    netFundingPnlValue: parseNumber(row["Net Funding P&L"]),
    netFundingReturnPct: row["Net Funding Return %"] || "",
    netFundingReturnPctValue: parsePercent(row["Net Funding Return %"] || ""),
  };
}

function buildRecommendationDateBySymbol(reportRows) {
  return reportRows.reduce((mapping, row) => {
    const stockCode = row["Stock Code/Name"] || "";
    const recommendationDate = row["Recommendation Date"] || "";
    if (!stockCode || !recommendationDate) {
      return mapping;
    }
    mapping[normalizeSymbol(stockCode)] = recommendationDate;
    return mapping;
  }, {});
}

function findInvestment(reportStockCode) {
  const normalized = normalizeSymbol(reportStockCode);
  const matching = pnlRows.filter(
    (item) =>
      (
        normalizeSymbol(item.matchedReportSymbol) === normalized ||
        normalizeSymbol(item.stockCode) === normalized
      )
  );
  if (matching.length === 0) {
    return null;
  }

  const totalBuyValue = matching.reduce((sum, item) => sum + item.buyValueRaw, 0);
  const totalPnlValue = matching.reduce((sum, item) => sum + item.totalPnlValue, 0);
  const totalReturnPctValue = totalBuyValue ? (totalPnlValue / totalBuyValue) * 100 : 0;

  return {
    buyValue: formatCurrency(totalBuyValue),
    returnPct: `${totalReturnPctValue.toFixed(2)}%`,
    returnPctValue: totalReturnPctValue,
  };
}

function findPositionStatus(reportStockCode) {
  const normalized = normalizeSymbol(reportStockCode);
  const matching = pnlRows.filter(
    (item) =>
      (
        normalizeSymbol(item.matchedReportSymbol) === normalized ||
        normalizeSymbol(item.stockCode) === normalized
      )
  );

  if (matching.length === 0) {
    return "No Position";
  }

  const totalNetQuantity = matching.reduce((sum, item) => sum + item.netQuantityValue, 0);
  const totalSellQuantity = matching.reduce((sum, item) => sum + item.sellQuantityValue, 0);
  if (totalNetQuantity <= 0 && totalSellQuantity > 0) {
    return "Realized";
  }
  if (totalSellQuantity > 0 && totalNetQuantity > 0) {
    return "Partially Realized";
  }
  return "Unrealized";
}

function parsePercent(value) {
  return Number.parseFloat(String(value).replace("%", "")) || 0;
}

function parseNumber(value) {
  return Number.parseFloat(String(value).replace(/,/g, "")) || 0;
}

function parseOptionalNumber(value) {
  const cleaned = String(value ?? "").replace(/,/g, "").trim();
  if (!cleaned) {
    return null;
  }
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeSymbol(value) {
  return String(value).trim().toUpperCase().replace(/\.NS$|\.BO$/g, "");
}

function distributeTargets(firstTarget, secondTarget) {
  if (firstTarget === null && secondTarget === null) {
    return [null, null, null, null];
  }
  if (firstTarget === null) {
    return [secondTarget, secondTarget, secondTarget, secondTarget];
  }
  if (secondTarget === null) {
    return [firstTarget, firstTarget, firstTarget, firstTarget];
  }

  const step = (secondTarget - firstTarget) / 3;
  const target2 = Math.round(firstTarget + step);
  const target3 = Math.round(firstTarget + (2 * step));
  return [firstTarget, target2, target3, secondTarget];
}

function classifyTargetHit(highestPrice, target) {
  if (!Number.isFinite(target)) {
    return "";
  }
  if (!Number.isFinite(highestPrice)) {
    return "FALSE";
  }
  return highestPrice >= target ? "TRUE" : "FALSE";
}

function formatTarget(value) {
  if (!Number.isFinite(value)) {
    return "";
  }
  return value.toFixed(2);
}

function formatTargetReturnPct(targetValue, buyPrice) {
  if (!Number.isFinite(targetValue) || !Number.isFinite(buyPrice) || buyPrice === 0) {
    return "";
  }
  const value = ((targetValue / buyPrice) - 1) * 100;
  return `${value.toFixed(2)}%`;
}

function formatPercent(value) {
  const sign = value > 0 ? "+" : "";
  return `${sign}${value.toFixed(2)}%`;
}

function formatCurrency(value) {
  const sign = value > 0 ? "+" : "";
  return `${sign}${value.toFixed(2)}`;
}

function renderTargetStatus(hitValue, amount, returnPct) {
  if (!hitValue) {
    return "--";
  }

  const isHit = hitValue === "TRUE";
  const symbol = isHit ? "✓" : "✗";
  const klass = isHit ? "hit" : "miss";
  const detailParts = [];
  if (amount) {
    detailParts.push(`Amount: ${amount}`);
  }
  if (returnPct) {
    detailParts.push(`% return: ${returnPct}`);
  }
  const detail = detailParts.length ? ` <span class="target-subdetail">(${detailParts.join(", ")})</span>` : "";

  return `<span class="target-status ${klass}">${symbol}</span><span class="target-detail">${detail}</span>`;
}

function updatePnlSummary(items) {
  const realized = items.reduce((sum, item) => sum + item.realizedPnlValue, 0);
  const charges = items.reduce((sum, item) => sum + item.chargesTaxesOthersValue, 0);
  const netRealized = items.reduce((sum, item) => sum + item.netRealizedPnlValue, 0);
  const unrealized = items.reduce((sum, item) => sum + item.unrealizedPnlValue, 0);
  const total = items.reduce((sum, item) => sum + item.totalPnlValue, 0);
  const investedBase = items.reduce((sum, item) => sum + getPnlRowBaseValue(item), 0);

  setSummaryValue(totalRealizedPnl, realized, investedBase, items.length > 0);
  totalChargesTaxesOthers.classList.remove("table-positive", "table-negative");
  totalChargesTaxesOthers.textContent = items.length > 0 ? formatCurrency(charges) : "--";
  if (items.length > 0 && charges > 0) {
    totalChargesTaxesOthers.classList.add("table-negative");
  }
  setSummaryValue(netRealizedPnl, netRealized, investedBase, items.length > 0);
  setSummaryValue(totalUnrealizedPnl, unrealized, investedBase, items.length > 0);
  totalOtherCreditsDebits.classList.remove("table-positive", "table-negative");
  totalOtherCreditsDebits.textContent = items.length > 0 ? formatCurrency(otherCreditsDebitsValue) : "--";
  if (items.length > 0) {
    if (otherCreditsDebitsValue > 0) {
      totalOtherCreditsDebits.classList.add("table-positive");
    } else if (otherCreditsDebitsValue < 0) {
      totalOtherCreditsDebits.classList.add("table-negative");
    }
  }
  setSummaryValue(netTotalPnl, total + otherCreditsDebitsValue, investedBase, items.length > 0);
  totalAmountInvested.classList.remove("table-positive", "table-negative");
  totalAmountInvested.textContent = items.length > 0 ? formatCurrency(investedBase) : "--";
}

function computeOtherCreditsDebits(rows) {
  return rows.reduce((sum, row) => {
    const particulars = String(row.particulars || "").trim().toLowerCase();
    if (!particulars || particulars === "opening balance" || particulars === "closing balance") {
      return sum;
    }
    if (
      particulars.includes("funds added using upi") ||
      particulars.includes("net settlement for equity") ||
      particulars.includes("net obligation for equity") ||
      particulars.includes("initial margin charged for mtf") ||
      particulars.includes("mtm obligation blocked for mtf") ||
      particulars.includes("mtm obligation reversed for mtf")
    ) {
      return sum;
    }
    const debit = parseNumber(row.debit);
    const credit = parseNumber(row.credit);
    return sum + (credit - debit);
  }, 0);
}

function getPnlRowBaseValue(item) {
  if (activePnlBasis === PNL_BASIS.PERSONAL_FUNDING) {
    return item.myFundingRaw || 0;
  }
  return item.buyValueRaw || 0;
}

function setSummaryValue(element, value, base, hasItems) {
  element.classList.remove("table-positive", "table-negative");

  if (!hasItems) {
    element.textContent = "--";
    return;
  }

  const pct = base ? (value / base) * 100 : 0;
  element.innerHTML = `${formatCurrency(value)} <span class="summary-percent">(${formatPercent(pct)})</span>`;

  if (value > 0) {
    element.classList.add("table-positive");
    return;
  }
  if (value < 0) {
    element.classList.add("table-negative");
  }
}

function updateMtfSummary(items) {
  const totalHolding = items.reduce((sum, item) => sum + item.holdingPnlValue, 0);
  const totalNetFundingPnl = items.reduce((sum, item) => sum + item.netFundingPnlValue, 0);
  const totalFundingBase = items.reduce((sum, item) => sum + item.yourFundingRaw, 0);
  const totalHoldingBase = items.reduce((sum, item) => sum + item.buyValueRaw, 0);
  const totalCarryingCost = items.reduce((sum, item) => sum + item.totalCarryingCostValue, 0);

  setSummaryValue(mtfHoldingPnl, totalHolding, totalHoldingBase, items.length > 0);
  setSummaryValue(mtfNetFundingPnl, totalNetFundingPnl, totalFundingBase, items.length > 0);
  setSummaryValue(mtfTotalCarryingCost, -totalCarryingCost, totalFundingBase, items.length > 0);
  mtfYourFunding.classList.remove("table-positive", "table-negative");
  mtfYourFunding.textContent = items.length > 0 ? formatCurrency(totalFundingBase) : "--";
}

function renderInvestmentTimeline(filteredPnlRows) {
  if (portfolioChart) {
    portfolioChart.destroy();
    portfolioChart = null;
  }

  if (!filteredPnlRows.length) {
    timelineTotalAdded.textContent = "--";
    timelineTotalWithdrawn.textContent = "--";
    timelineCurrentGain.textContent = "--";
    timelineCurrentGain.classList.remove("table-positive", "table-negative");
    return;
  }

  const capitalKey = activePnlBasis === PNL_BASIS.PERSONAL_FUNDING ? "myFundingRaw" : "buyValueRaw";
  const totalInvested = filteredPnlRows.reduce((s, r) => s + r[capitalKey], 0);
  const totalSold = filteredPnlRows.reduce((s, r) => s + r.sellValueRaw, 0);
  const totalPnl = filteredPnlRows.reduce((s, r) => s + r.totalPnlValue, 0) + otherCreditsDebitsValue;

  timelineTotalAdded.textContent = formatCurrency(totalInvested);
  timelineTotalWithdrawn.textContent = formatCurrency(totalSold);
  timelineCurrentGain.classList.remove("table-positive", "table-negative");
  timelineCurrentGain.textContent = formatCurrency(totalPnl);
  if (totalPnl > 0) {
    timelineCurrentGain.classList.add("table-positive");
  } else if (totalPnl < 0) {
    timelineCurrentGain.classList.add("table-negative");
  }

  const timeline = buildPortfolioTimeline(filteredPnlRows, capitalKey, otherCreditsDebitsValue);
  if (timeline.length === 0) {
    return;
  }

  const canvas = document.getElementById("investmentTimelineChart");
  portfolioChart = new Chart(canvas, {
    type: "line",
    data: {
      labels: timeline.map((p) => p.date),
      datasets: [
        {
          label: "Input Capital",
          data: timeline.map((p) => p.inputCapital),
          borderColor: "#5b8dee",
          backgroundColor: "rgba(91,141,238,0.08)",
          borderWidth: 2.4,
          borderDash: [7, 4],
          tension: 0.3,
          pointRadius: 4,
          pointHoverRadius: 6,
          fill: false,
        },
        {
          label: "Portfolio Value",
          data: timeline.map((p) => p.portfolioValue),
          borderColor: "#00e5a0",
          backgroundColor: "rgba(0,229,160,0.1)",
          borderWidth: 3.2,
          tension: 0.3,
          pointRadius: 4,
          pointHoverRadius: 6,
          fill: false,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false,
      },
      plugins: {
        tooltip: {
          backgroundColor: "rgba(7,12,26,0.92)",
          borderColor: "rgba(255,255,255,0.1)",
          borderWidth: 1,
          titleColor: "#dde4f5",
          bodyColor: "#5a6b8a",
          callbacks: {
            label: (ctx) => ` ${ctx.dataset.label}: ${formatCompactCurrency(ctx.parsed.y)}`,
          },
        },
        legend: {
          display: true,
          position: "top",
          labels: { font: { size: 12 }, color: "#5a6b8a" },
        },
      },
      scales: {
        x: {
          ticks: { maxTicksLimit: 10, color: "#5a6b8a", font: { size: 11 } },
          grid: { display: false },
        },
        y: {
          ticks: {
            callback: (v) => formatCompactCurrency(v),
            color: "#5a6b8a",
            font: { size: 11 },
          },
          grid: { color: "rgba(255,255,255,0.06)" },
        },
      },
    },
  });
}

function buildPortfolioTimeline(filteredPnlRows, capitalKey = "buyValueRaw", otherPnl = 0) {
  const totalCapital = filteredPnlRows.reduce((s, r) => s + r[capitalKey], 0);
  const totalPnl = filteredPnlRows.reduce((s, r) => s + r.totalPnlValue, 0) + otherPnl;
  const totalPortfolioValue = totalCapital + totalPnl;

  if (totalCapital <= 0) {
    return [];
  }

  const byDate = {};
  filteredPnlRows.forEach((row) => {
    const date = (row.tradeDate || "").trim();
    if (!date) {
      return;
    }
    if (!byDate[date]) {
      byDate[date] = 0;
    }
    byDate[date] += row[capitalKey];
  });

  const sorted = Object.keys(byDate).sort();
  if (sorted.length === 0) {
    return [];
  }

  // rows with no tradeDate are excluded from byDate — add their capital to the last
  // date so the chart's final inputCapital and portfolioValue match the summary stats exactly
  const byDateTotal = Object.values(byDate).reduce((s, v) => s + v, 0);
  const unattributed = totalCapital - byDateTotal;
  if (unattributed !== 0) {
    byDate[sorted[sorted.length - 1]] += unattributed;
  }

  const returnMultiplier = totalPortfolioValue / totalCapital;
  let cumCapital = 0;
  return sorted.map((date) => {
    cumCapital += byDate[date];
    return {
      date,
      inputCapital: cumCapital,
      portfolioValue: cumCapital * returnMultiplier,
    };
  });
}

function formatCompactCurrency(value) {
  const absolute = Math.abs(value);
  const sign = value < 0 ? "-" : "";
  if (absolute >= 10000000) {
    return `${sign}${(absolute / 10000000).toFixed(2)}Cr`;
  }
  if (absolute >= 100000) {
    return `${sign}${(absolute / 100000).toFixed(2)}L`;
  }
  if (absolute >= 1000) {
    return `${sign}${(absolute / 1000).toFixed(1)}K`;
  }
  return `${sign}${absolute.toFixed(0)}`;
}

function setPnlFilter(filter) {
  activePnlFilter = Object.values(STOCK_FILTER).includes(filter) ? filter : STOCK_FILTER.ALL;
  pnlFilterButtons.forEach((button) => {
    button.classList.toggle("active", (button.dataset.filter || STOCK_FILTER.ALL) === activePnlFilter);
  });
  refreshPnlView();
  renderInvestmentTimeline(getCurrentPnlRows());
}

function setPnlBasis(basis) {
  activePnlBasis = Object.values(PNL_BASIS).includes(basis) ? basis : PNL_BASIS.ALL_HOLDINGS;
  pnlBasisButtons.forEach((button) => {
    button.classList.toggle("active", (button.dataset.basis || PNL_BASIS.ALL_HOLDINGS) === activePnlBasis);
  });
  refreshPnlView();
  renderInvestmentTimeline(getCurrentPnlRows());
}

function getPnlRowsByActiveFilter(items) {
  if (activePnlFilter === STOCK_FILTER.REALIZED) {
    return items.filter((item) => item.realizedPnlValue !== 0);
  }
  if (activePnlFilter === STOCK_FILTER.UNREALIZED) {
    return items.filter((item) => item.unrealizedPnlValue !== 0);
  }
  return items;
}

function setMtfFilter(filter) {
  activeMtfFilter = Object.values(STOCK_FILTER).includes(filter) ? filter : STOCK_FILTER.ALL;
  mtfFilterButtons.forEach((button) => {
    button.classList.toggle("active", (button.dataset.filter || STOCK_FILTER.ALL) === activeMtfFilter);
  });
  refreshMtfView();
}

function getMtfRowsByActiveFilter(items) {
  if (activeMtfFilter === STOCK_FILTER.REALIZED) {
    return items.filter((item) => item.sellQuantityValue > 0);
  }
  if (activeMtfFilter === STOCK_FILTER.UNREALIZED) {
    return items.filter((item) => item.netQuantityValue > 0);
  }
  return items;
}

function setView(view) {
  activeView = view;
  const showingCharts = view === VIEW.CHARTS;
  const showingPnl = view === VIEW.PNL;
  const showingMtf = view === VIEW.MTF;
  chartViewButton.classList.toggle("active", showingCharts);
  pnlViewButton.classList.toggle("active", showingPnl);
  mtfViewButton.classList.toggle("active", showingMtf);
  chartsPage.classList.toggle("active", showingCharts);
  pnlPage.classList.toggle("active", showingPnl);
  mtfPage.classList.toggle("active", showingMtf);
  pnlFilterSwitch.classList.toggle("hidden", !showingPnl);
  pnlBasisSwitch.classList.toggle("hidden", !showingPnl);
  mtfFilterSwitch.classList.toggle("hidden", !showingMtf);

  if (showingCharts) {
    renderCards(searchInput.value);
    return;
  }

  if (showingPnl) {
    refreshPnlView();
    return;
  }

  refreshMtfView();
}

function getTableByKey(tableKey) {
  if (tableKey === TABLE_KEYS.PNL) {
    return pnlTable;
  }
  return mtfTable;
}

function getExcludedContainerByKey(tableKey) {
  if (tableKey === TABLE_KEYS.PNL) {
    return pnlExcludedColumns;
  }
  return mtfExcludedColumns;
}

function getHeaderCells(tableKey) {
  return Array.from(getTableByKey(tableKey).querySelectorAll("thead th"));
}

function getVisibleColumnCount(tableKey) {
  const totalColumns = getHeaderCells(tableKey).length;
  const hiddenColumns = hiddenColumnsByTable[tableKey].size;
  return Math.max(1, totalColumns - hiddenColumns);
}

function initializeColumnControls() {
  Object.values(TABLE_KEYS).forEach((tableKey) => {
    const headers = getHeaderCells(tableKey);
    headers.forEach((th, index) => {
      const label = th.textContent.trim();
      th.dataset.columnLabel = label;
      const wrapper = document.createElement("span");
      wrapper.className = "column-header-content";

      const labelNode = document.createElement("span");
      labelNode.textContent = label;
      wrapper.appendChild(labelNode);

      const button = document.createElement("button");
      button.type = "button";
      button.className = "column-hide-button";
      button.textContent = "x";
      button.title = `Hide ${label}`;
      button.addEventListener("click", () => {
        hiddenColumnsByTable[tableKey].add(index);
        applyColumnVisibility(tableKey);
        renderExcludedColumnPills(tableKey);
      });
      wrapper.appendChild(button);

      th.textContent = "";
      th.appendChild(wrapper);
    });
    applyColumnVisibility(tableKey);
    renderExcludedColumnPills(tableKey);
  });
}

function applyColumnVisibility(tableKey) {
  const table = getTableByKey(tableKey);
  const hiddenColumns = hiddenColumnsByTable[tableKey];
  const rows = table.querySelectorAll("tr");

  rows.forEach((row) => {
    Array.from(row.children).forEach((cell, index) => {
      cell.style.display = hiddenColumns.has(index) ? "none" : "";
    });
  });
}

function renderExcludedColumnPills(tableKey) {
  const container = getExcludedContainerByKey(tableKey);
  const headers = getHeaderCells(tableKey);
  const hiddenColumns = Array.from(hiddenColumnsByTable[tableKey]).sort((a, b) => a - b);
  container.innerHTML = "";

  hiddenColumns.forEach((columnIndex) => {
    const th = headers[columnIndex];
    if (!th) {
      return;
    }
    const pill = document.createElement("button");
    pill.type = "button";
    pill.className = "excluded-pill";
    pill.textContent = th.dataset.columnLabel || `Column ${columnIndex + 1}`;
    pill.title = `Show ${pill.textContent}`;
    pill.addEventListener("click", () => {
      hiddenColumnsByTable[tableKey].delete(columnIndex);
      applyColumnVisibility(tableKey);
      renderExcludedColumnPills(tableKey);
    });
    container.appendChild(pill);
  });
}

function setStatus(message, isError) {
  statusBanner.textContent = message;
  statusBanner.classList.toggle("hidden", !message);
  statusBanner.classList.toggle("error", isError);
}

function parseCsv(csvText) {
  const rows = [];
  let row = [];
  let value = "";
  let insideQuotes = false;

  for (let index = 0; index < csvText.length; index += 1) {
    const character = csvText[index];
    const nextCharacter = csvText[index + 1];

    if (character === '"') {
      if (insideQuotes && nextCharacter === '"') {
        value += '"';
        index += 1;
      } else {
        insideQuotes = !insideQuotes;
      }
      continue;
    }

    if (character === "," && !insideQuotes) {
      row.push(value);
      value = "";
      continue;
    }

    if ((character === "\n" || character === "\r") && !insideQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      continue;
    }

    value += character;
  }

  if (value.length > 0 || row.length > 0) {
    row.push(value);
    rows.push(row);
  }

  const [headerRow = [], ...dataRows] = rows.filter((entry) => entry.some((cell) => cell !== ""));
  return dataRows.map((dataRow) =>
    Object.fromEntries(headerRow.map((header, index) => [header, dataRow[index] || ""]))
  );
}
