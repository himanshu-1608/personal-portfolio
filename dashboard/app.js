const REPORT_PATH = "../reports/stock_closing_prices.csv";
const PNL_REPORT_PATH = "../reports/stock_pnl_summary.csv";
const MTF_PNL_REPORT_PATH = "../reports/mtf_pnl_summary.csv";
const LEDGER_PATH = "../reports/exported_all_ledger.csv";
const PORTFOLIO_TIMELINE_PATH = "../reports/portfolio_timeline.csv";
const DAY_COUNT = 252;
const IS_FILE_PROTOCOL = window.location.protocol === "file:";
const VIEW = {
  CHARTS: "charts",
  PNL: "pnl",
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
};

// Shareable URL state. ?view=<tab>&filter=<realized|unrealized|all>&basis=<total|mymoney>
const VIEW_TO_PARAM = {
  [VIEW.CHARTS]: "recommendations",
  [VIEW.PNL]: "portfolio",
};
const PARAM_TO_VIEW = {
  recommendations: VIEW.CHARTS,
  portfolio: VIEW.PNL,
};
const BASIS_TO_PARAM = {
  [PNL_BASIS.ALL_HOLDINGS]: "total",
  [PNL_BASIS.PERSONAL_FUNDING]: "mymoney",
};
const PARAM_TO_BASIS = {
  total: PNL_BASIS.ALL_HOLDINGS,
  mymoney: PNL_BASIS.PERSONAL_FUNDING,
};

const searchInput = document.getElementById("searchInput");
const statusBanner = document.getElementById("statusBanner");
const cardsContainer = document.getElementById("cardsContainer");
const template = document.getElementById("stockCardTemplate");

const pnlFilterSwitch = document.getElementById("pnlFilterSwitch");
const pnlFilterButtons = Array.from(pnlFilterSwitch.querySelectorAll(".pnl-filter-button"));
const pnlBasisSwitch = document.getElementById("pnlBasisSwitch");
const pnlBasisButtons = Array.from(pnlBasisSwitch.querySelectorAll(".pnl-basis-button"));
const chartViewButton = document.getElementById("chartViewButton");
const pnlViewButton = document.getElementById("pnlViewButton");
const chartsPage = document.getElementById("chartsPage");
const pnlPage = document.getElementById("pnlPage");
const pnlTable = document.getElementById("pnlTable");
const pnlTableBody = document.getElementById("pnlTableBody");
const pnlExcludedColumns = document.getElementById("pnlExcludedColumns");
const totalRealizedPnl = document.getElementById("totalRealizedPnl");
const totalChargesTaxesOthers = document.getElementById("totalChargesTaxesOthers");
const netRealizedPnl = document.getElementById("netRealizedPnl");
const totalUnrealizedPnl = document.getElementById("totalUnrealizedPnl");
const totalOtherCreditsDebits = document.getElementById("totalOtherCreditsDebits");
const netTotalPnl = document.getElementById("netTotalPnl");
const totalAmountInvested = document.getElementById("totalAmountInvested");
const timelineTotalAdded = document.getElementById("timelineTotalAdded");
const timelineTotalWithdrawn = document.getElementById("timelineTotalWithdrawn");
const timelineCurrentGain = document.getElementById("timelineCurrentGain");

let allStocks = [];
let pnlRows = [];
let mtfRows = [];
let ledgerRows = [];
// "Other Credits & Debits" segregated so each filter combo only sees its share.
// realizedCarrying / unrealizedCarrying: MTF interest + pledge costs (a leverage
// cost, so they belong to "My Money" basis only), split by whether the symbol's
// position is closed or still held. realizedDp: DP charges on sales (both bases,
// realized only). Account-level charges (DDPI, bank, margin blocks) are out of scope.
let otherCreditsBreakdown = { realizedCarrying: 0, unrealizedCarrying: 0, realizedDp: 0 };
let portfolioChart = null;
let cardCharts = [];
let realPortfolioTimeline = [];
let activeView = VIEW.CHARTS;
let activePnlFilter = STOCK_FILTER.ALL;
let activePnlBasis = PNL_BASIS.ALL_HOLDINGS;
const hiddenColumnsByTable = {
  [TABLE_KEYS.PNL]: new Set(),
};

searchInput.addEventListener("input", () => {
  if (activeView === VIEW.CHARTS) {
    renderCards(searchInput.value);
  } else {
    refreshPnlView();
  }
});

chartViewButton.addEventListener("click", () => {
  setView(VIEW.CHARTS);
});

pnlViewButton.addEventListener("click", () => {
  setView(VIEW.PNL);
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

loadDashboard();
initializeColumnControls();

async function loadDashboard() {
  setStatus("", false);

  try {
    const { reportCsvText, pnlCsvText, mtfCsvText, ledgerCsvText, portfolioTimelineCsvText } = await loadCsvSources();
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
    otherCreditsBreakdown = computeOtherCreditsBreakdown(ledgerRows, mtfRows);
    realPortfolioTimeline = portfolioTimelineCsvText ? parseCsv(portfolioTimelineCsvText) : [];

    updateSummary(allStocks);
    renderCards(searchInput.value);
    refreshPnlView();
    applyStateFromUrl();

    if (allStocks.length === 0) {
      setStatus("The report has no chartable rows yet. Run python/fetch_yahoo_prices.py to refresh CSV reports.", false);
    }
  } catch (error) {
    allStocks = [];
    pnlRows = [];
    mtfRows = [];
    ledgerRows = [];
    otherCreditsBreakdown = { realizedCarrying: 0, unrealizedCarrying: 0, realizedDp: 0 };
    updateSummary([]);
    updatePnlSummary([]);
    cardsContainer.innerHTML = "";
    pnlTableBody.innerHTML = "";
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

function refreshPnlView() {
  updatePnlSummary(getCurrentPnlRows());
  renderPnlTable(searchInput.value);
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
    portfolioTimelineCsvText: embeddedData.portfolioTimelineCsv,
  };
}

async function fetchCsvSources() {
  const cacheBuster = Date.now();
  const [reportResponse, pnlResponse, mtfResponse, ledgerResponse, timelineResponse] = await Promise.all([
    fetch(`${REPORT_PATH}?t=${cacheBuster}`),
    fetch(`${PNL_REPORT_PATH}?t=${cacheBuster}`),
    fetch(`${MTF_PNL_REPORT_PATH}?t=${cacheBuster}`),
    fetch(`${LEDGER_PATH}?t=${cacheBuster}`),
    fetch(`${PORTFOLIO_TIMELINE_PATH}?t=${cacheBuster}`),
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
  const portfolioTimelineCsvText = timelineResponse.ok ? await timelineResponse.text() : "";
  return { reportCsvText, pnlCsvText, mtfCsvText, ledgerCsvText, portfolioTimelineCsvText };
}

function readEmbeddedData() {
  const data = window.__DASHBOARD_DATA__;
  if (!data || typeof data !== "object") {
    return {
      stockClosingPricesCsv: "",
      stockPnlSummaryCsv: "",
      mtfPnlSummaryCsv: "",
      allLedgerCsv: "",
      portfolioTimelineCsv: "",
    };
  }

  return {
    stockClosingPricesCsv: String(data.stockClosingPricesCsv || ""),
    stockPnlSummaryCsv: String(data.stockPnlSummaryCsv || ""),
    mtfPnlSummaryCsv: String(data.mtfPnlSummaryCsv || ""),
    allLedgerCsv: String(data.allLedgerCsv || ""),
    portfolioTimelineCsv: String(data.portfolioTimelineCsv || ""),
  };
}

function updateSummary(_items) {}

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

  cardCharts.forEach((chart) => chart.destroy());
  cardCharts = [];
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

    fillTable(fragment.querySelector(".data-table-body"), item.points);

    card.style.animationDelay = `${index * 70}ms`;
    cardsContainer.appendChild(fragment);
    // Chart.js must measure a canvas that is already in the DOM, so draw after append.
    drawChart(card.querySelector(".chart"), item.points);
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
      <td>${item.stockCode}${item.hasMtfFunding ? ' <span class="mtf-tag">MTF</span>' : ""}</td>
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

// Interactive return-% chart for each recommendation card. Mirrors the portfolio
// timeline's Chart.js interactivity: hover crosshair + tooltip (date / price /
// return), sign-aware colouring, and a gradient area fill. Chart instances are
// tracked in cardCharts and destroyed before each re-render.
function drawChart(canvas, points) {
  if (!canvas || typeof Chart === "undefined" || points.length === 0) {
    return;
  }

  const labels = points.map((point) => point.date.slice(5));
  const values = points.map((point) => point.returnValue);
  const positive = (values[values.length - 1] ?? 0) >= 0;
  const lineColor = positive ? "#00e5a0" : "#ff4d6a";
  const fillTop = positive ? "rgba(0,229,160,0.28)" : "rgba(255,77,106,0.26)";

  // Dashed zero baseline + a vertical crosshair under the hovered point.
  const zeroAndCrosshair = {
    id: "zeroAndCrosshair",
    afterDatasetsDraw(chart) {
      const { ctx, chartArea, scales } = chart;
      if (!chartArea) {
        return;
      }
      if (scales.y.min < 0 && scales.y.max > 0) {
        const y = scales.y.getPixelForValue(0);
        ctx.save();
        ctx.beginPath();
        ctx.setLineDash([4, 5]);
        ctx.moveTo(chartArea.left, y);
        ctx.lineTo(chartArea.right, y);
        ctx.strokeStyle = "rgba(255,255,255,0.2)";
        ctx.lineWidth = 1;
        ctx.stroke();
        ctx.restore();
      }
      const active = chart.tooltip ? chart.tooltip.getActiveElements() : [];
      if (active.length) {
        const x = active[0].element.x;
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(x, chartArea.top);
        ctx.lineTo(x, chartArea.bottom);
        ctx.strokeStyle = "rgba(255,255,255,0.24)";
        ctx.lineWidth = 1;
        ctx.stroke();
        ctx.restore();
      }
    },
  };

  const chart = new Chart(canvas, {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          data: values,
          borderColor: lineColor,
          borderWidth: 2.6,
          tension: 0.35,
          fill: true,
          backgroundColor: (context) => {
            const { chart: c } = context;
            if (!c.chartArea) {
              return fillTop;
            }
            const gradient = c.ctx.createLinearGradient(0, c.chartArea.top, 0, c.chartArea.bottom);
            gradient.addColorStop(0, fillTop);
            gradient.addColorStop(1, "rgba(0,0,0,0)");
            return gradient;
          },
          pointRadius: 0,
          pointHoverRadius: 6,
          pointHoverBackgroundColor: lineColor,
          pointHoverBorderColor: "#0b1220",
          pointHoverBorderWidth: 2,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: "rgba(10,16,28,0.94)",
          borderColor: "rgba(255,255,255,0.1)",
          borderWidth: 1,
          padding: 11,
          cornerRadius: 10,
          titleColor: "#dde4f5",
          bodyColor: "#aeb9d6",
          displayColors: false,
          callbacks: {
            title: (items) => points[items[0].dataIndex].date,
            label: (item) => {
              const point = points[item.dataIndex];
              return [`Price  ${point.price}`, `Return  ${point.returnText}`];
            },
          },
        },
      },
      scales: {
        y: {
          grid: { color: "rgba(255,255,255,0.06)" },
          border: { display: false },
          ticks: {
            color: "#8d9bb8",
            font: { size: 11 },
            maxTicksLimit: 6,
            callback: (value) => `${value}%`,
          },
        },
        x: {
          grid: { display: false },
          border: { color: "rgba(255,255,255,0.1)" },
          ticks: { color: "#8d9bb8", font: { size: 11 }, maxTicksLimit: 6, autoSkip: true },
        },
      },
      animation: { duration: 520, easing: "easeOutCubic" },
    },
    plugins: [zeroAndCrosshair],
  });

  cardCharts.push(chart);
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
  const highestPriceValue = parseNumber(row["Highest Price"]);

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
    highestPrice: row["Highest Price"] || "",
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
  const otherCredits = otherCreditsForFilters(activePnlFilter, activePnlBasis);
  totalOtherCreditsDebits.classList.remove("table-positive", "table-negative");
  totalOtherCreditsDebits.textContent = items.length > 0 ? formatCurrency(otherCredits) : "--";
  if (items.length > 0) {
    if (otherCredits > 0) {
      totalOtherCreditsDebits.classList.add("table-positive");
    } else if (otherCredits < 0) {
      totalOtherCreditsDebits.classList.add("table-negative");
    }
  }
  setSummaryValue(netTotalPnl, total + otherCredits, investedBase, items.length > 0);
  totalAmountInvested.classList.remove("table-positive", "table-negative");
  totalAmountInvested.textContent = items.length > 0 ? formatCurrency(investedBase) : "--";

  // Hide the opposite summary card: Realized filter drops the Unrealized card and vice versa.
  const realizedCard = totalRealizedPnl.closest(".summary-card");
  const unrealizedCard = totalUnrealizedPnl.closest(".summary-card");
  if (realizedCard) {
    realizedCard.classList.toggle("hidden", activePnlFilter === STOCK_FILTER.UNREALIZED);
  }
  if (unrealizedCard) {
    unrealizedCard.classList.toggle("hidden", activePnlFilter === STOCK_FILTER.REALIZED);
  }
}

// Build the segregated breakdown from the MTF lots (carrying cost + open/closed
// status) and the ledger (DP sale charges). Costs are stored as negatives.
function computeOtherCreditsBreakdown(ledger, mtf) {
  let realizedCarrying = 0;
  let unrealizedCarrying = 0;
  for (const lot of mtf) {
    const carrying = Math.abs(lot.totalCarryingCostValue || 0);
    if (carrying === 0) {
      continue;
    }
    if (lot.netQuantityValue > 0) {
      unrealizedCarrying -= carrying; // still held → unrealized
    } else {
      realizedCarrying -= carrying; // fully sold → realized
    }
  }

  let realizedDp = 0;
  for (const row of ledger) {
    const particulars = String(row.particulars || "").trim().toLowerCase();
    if (particulars.includes("dp charges for sale")) {
      realizedDp -= parseNumber(row.debit); // sale event → realized, both bases
    }
  }

  return { realizedCarrying, unrealizedCarrying, realizedDp };
}

// Amount of "Other Credits & Debits" applicable to the active filter + basis.
// Carrying costs (leverage costs) only count under "My Money" (personal_funding).
function otherCreditsForFilters(filter, basis) {
  const b = otherCreditsBreakdown;
  const personal = basis === PNL_BASIS.PERSONAL_FUNDING;
  if (filter === STOCK_FILTER.REALIZED) {
    return b.realizedDp + (personal ? b.realizedCarrying : 0);
  }
  if (filter === STOCK_FILTER.UNREALIZED) {
    return personal ? b.unrealizedCarrying : 0;
  }
  return b.realizedDp + (personal ? b.realizedCarrying + b.unrealizedCarrying : 0);
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


// Derive one {date, inputCapital, portfolioValue} point from a timeline CSV row,
// honoring the two active filters:
//   activePnlBasis  -> all_holdings (full value) | personal_funding (full - zerodha)
//   activePnlFilter -> all | realized | unrealized (which bucket to include)
function timelinePointForFilters(row) {
  const rInvFull = parseNumber(row.r_invested);
  const rValFull = parseNumber(row.r_value);
  const rZerodha = parseNumber(row.r_zerodha);
  const uInvFull = parseNumber(row.u_invested);
  const uValFull = parseNumber(row.u_value);
  const uZerodha = parseNumber(row.u_zerodha);

  const personal = activePnlBasis === PNL_BASIS.PERSONAL_FUNDING;
  const rInv = personal ? rInvFull - rZerodha : rInvFull;
  const rVal = personal ? rValFull - rZerodha : rValFull;
  const uInv = personal ? uInvFull - uZerodha : uInvFull;
  const uVal = personal ? uValFull - uZerodha : uValFull;

  let inputCapital;
  let portfolioValue;
  if (activePnlFilter === STOCK_FILTER.REALIZED) {
    inputCapital = rInv;
    portfolioValue = rVal;
  } else if (activePnlFilter === STOCK_FILTER.UNREALIZED) {
    inputCapital = uInv;
    portfolioValue = uVal;
  } else {
    inputCapital = rInv + uInv;
    portfolioValue = rVal + uVal;
  }

  return { date: row.date, inputCapital, portfolioValue };
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
  const otherCredits = otherCreditsForFilters(activePnlFilter, activePnlBasis);
  const totalPnl = filteredPnlRows.reduce((s, r) => s + r.totalPnlValue, 0) + otherCredits;

  timelineTotalAdded.textContent = formatCurrency(totalInvested);
  timelineTotalWithdrawn.textContent = formatCurrency(totalSold);
  timelineCurrentGain.classList.remove("table-positive", "table-negative");
  timelineCurrentGain.textContent = formatCurrency(totalPnl);
  if (totalPnl > 0) {
    timelineCurrentGain.classList.add("table-positive");
  } else if (totalPnl < 0) {
    timelineCurrentGain.classList.add("table-negative");
  }

  const timeline = realPortfolioTimeline.length > 0
    ? realPortfolioTimeline.map((row) => timelinePointForFilters(row))
    : buildPortfolioTimeline(filteredPnlRows, capitalKey, otherCredits);

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
          stepped: true,
          tension: 0,
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
          stepped: true,
          tension: 0,
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
  updateUrlParams();
}

function setPnlBasis(basis) {
  activePnlBasis = Object.values(PNL_BASIS).includes(basis) ? basis : PNL_BASIS.ALL_HOLDINGS;
  pnlBasisButtons.forEach((button) => {
    button.classList.toggle("active", (button.dataset.basis || PNL_BASIS.ALL_HOLDINGS) === activePnlBasis);
  });
  refreshPnlView();
  renderInvestmentTimeline(getCurrentPnlRows());
  updateUrlParams();
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

// Write the current tab + both filters into the URL (no reload) so the link is shareable.
function updateUrlParams() {
  try {
    const params = new URLSearchParams(window.location.search);
    params.set("view", VIEW_TO_PARAM[activeView] || "recommendations");
    params.set("filter", activePnlFilter);
    params.set("basis", BASIS_TO_PARAM[activePnlBasis] || "total");
    const newUrl = `${window.location.pathname}?${params.toString()}${window.location.hash}`;
    window.history.replaceState(null, "", newUrl);
  } catch (error) {
    /* file:// or unsupported history API — URL syncing is a no-op */
  }
}

// On load, restore tab + filters from the URL so a shared link opens the same state.
function applyStateFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const filterParam = (params.get("filter") || "").toLowerCase();
  const basisParam = (params.get("basis") || "").toLowerCase();
  const viewParam = (params.get("view") || "").toLowerCase();

  if (Object.values(STOCK_FILTER).includes(filterParam)) {
    activePnlFilter = filterParam;
  }
  if (PARAM_TO_BASIS[basisParam]) {
    activePnlBasis = PARAM_TO_BASIS[basisParam];
  }

  pnlFilterButtons.forEach((b) => b.classList.toggle("active", (b.dataset.filter || STOCK_FILTER.ALL) === activePnlFilter));
  pnlBasisButtons.forEach((b) => b.classList.toggle("active", (b.dataset.basis || PNL_BASIS.ALL_HOLDINGS) === activePnlBasis));

  setView(PARAM_TO_VIEW[viewParam] || VIEW.CHARTS);
  renderInvestmentTimeline(getCurrentPnlRows());
}

function setView(view) {
  activeView = view;
  const showingCharts = view === VIEW.CHARTS;
  const showingPnl = view === VIEW.PNL;
  chartViewButton.classList.toggle("active", showingCharts);
  pnlViewButton.classList.toggle("active", showingPnl);
  chartsPage.classList.toggle("active", showingCharts);
  pnlPage.classList.toggle("active", showingPnl);
  pnlFilterSwitch.classList.toggle("hidden", !showingPnl);
  pnlBasisSwitch.classList.toggle("hidden", !showingPnl);

  if (showingCharts) {
    renderCards(searchInput.value);
  } else {
    refreshPnlView();
  }

  updateUrlParams();
}

function getTableByKey(_tableKey) {
  return pnlTable;
}

function getExcludedContainerByKey(_tableKey) {
  return pnlExcludedColumns;
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
