const columns = {
  date: ["datum", "date", "entry date", "einstieg datum", "open date", "opened"],
  exitDate: ["ausstieg datum", "exit date", "close date", "closed"],
  symbol: ["symbol", "ticker", "instrument", "asset"],
  market: ["markt", "market", "asset class"],
  direction: ["richtung", "direction", "side", "long/short", "typ"],
  strategy: ["strategie", "strategy", "setup", "system"],
  entry: ["entry", "entry price", "einstieg", "einstiegspreis"],
  exit: ["exit", "exit price", "ausstieg", "ausstiegspreis"],
  quantity: ["menge", "quantity", "qty", "size", "position size"],
  risk: ["risiko", "risk", "risk eur", "risk €", "planned risk"],
  pnl: ["p&l", "pnl", "profit", "profit/loss", "gewinn", "ergebnis", "brutto"],
  fees: ["gebühren", "fees", "commission", "commissions", "kosten"],
  netPnl: ["net p&l", "net pnl", "p&l netto", "pnl netto", "netto"],
};

let trades = [];
let charts = {};

const euro = new Intl.NumberFormat("de-DE", {
  style: "currency",
  currency: "EUR",
  maximumFractionDigits: 2,
});
const number = new Intl.NumberFormat("de-DE", { maximumFractionDigits: 2 });

document.getElementById("excelFile").addEventListener("change", handleFile);
document.getElementById("loadDemo").addEventListener("click", loadDemo);
document.getElementById("clearData").addEventListener("click", clearData);
["dateFrom", "dateTo", "symbolFilter", "strategyFilter", "directionFilter", "resultFilter"].forEach((id) => {
  document.getElementById(id).addEventListener("change", render);
});

initSelects();
render();

async function handleFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  trades = rows.map(normalizeTrade).filter((trade) => trade.date || trade.symbol || trade.netPnl !== 0);
  document.getElementById("dataStatus").textContent = `${file.name} geladen`;
  setupFilters();
  render();
}

function normalizeTrade(row, index) {
  const normalized = {};
  Object.entries(row).forEach(([key, value]) => {
    normalized[String(key).trim().toLowerCase()] = value;
  });

  const rawPnl = pick(normalized, "pnl");
  const rawFees = pick(normalized, "fees");
  const net = pick(normalized, "netPnl");
  const pnl = toNumber(rawPnl);
  const fees = toNumber(rawFees);
  const netPnl = net === "" || net === undefined ? pnl - fees : toNumber(net);
  const risk = Math.abs(toNumber(pick(normalized, "risk")));

  return {
    id: index + 1,
    date: toDate(pick(normalized, "date") || pick(normalized, "exitDate")),
    symbol: clean(pick(normalized, "symbol")) || "Unbekannt",
    market: clean(pick(normalized, "market")),
    direction: normalizeDirection(pick(normalized, "direction")),
    strategy: clean(pick(normalized, "strategy")) || "Ohne Strategie",
    entry: toNumber(pick(normalized, "entry")),
    exit: toNumber(pick(normalized, "exit")),
    quantity: toNumber(pick(normalized, "quantity")),
    risk,
    pnl,
    fees,
    netPnl,
    rMultiple: risk > 0 ? netPnl / risk : 0,
  };
}

function pick(row, field) {
  const names = columns[field];
  const direct = names.find((name) => Object.prototype.hasOwnProperty.call(row, name));
  if (direct) return row[direct];

  const fuzzy = Object.keys(row).find((key) => names.some((name) => key.includes(name)));
  return fuzzy ? row[fuzzy] : "";
}

function toNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value ?? "")
    .replace(/\s/g, "")
    .replace(/[€$]/g, "")
    .replace(/\./g, "")
    .replace(",", ".");
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function toDate(value) {
  if (!value) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  const match = String(value).match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (!match) return "";
  const year = match[3].length === 2 ? `20${match[3]}` : match[3];
  return `${year}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
}

function clean(value) {
  return String(value ?? "").trim();
}

function normalizeDirection(value) {
  const text = clean(value).toLowerCase();
  if (text.includes("short") || text === "s" || text.includes("sell")) return "Short";
  if (text.includes("long") || text === "l" || text.includes("buy")) return "Long";
  return clean(value) || "Unbekannt";
}

function setupFilters() {
  fillSelect("symbolFilter", unique(trades.map((trade) => trade.symbol)));
  fillSelect("strategyFilter", unique(trades.map((trade) => trade.strategy)));
  fillSelect("directionFilter", unique(trades.map((trade) => trade.direction)));

  const dates = trades.map((trade) => trade.date).filter(Boolean).sort();
  document.getElementById("dateFrom").value = dates[0] || "";
  document.getElementById("dateTo").value = dates[dates.length - 1] || "";
}

function initSelects() {
  fillSelect("symbolFilter", []);
  fillSelect("strategyFilter", []);
  fillSelect("directionFilter", []);
}

function resetFilters() {
  document.getElementById("dateFrom").value = "";
  document.getElementById("dateTo").value = "";
  document.getElementById("resultFilter").value = "all";
  initSelects();
}

function fillSelect(id, values) {
  const select = document.getElementById(id);
  select.innerHTML = `<option value="all">Alle</option>${values
    .map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
    .join("")}`;
}

function unique(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => a.localeCompare(b, "de"));
}

function render() {
  const filtered = getFilteredTrades();
  renderKpis(filtered);
  renderTable(filtered);
  renderCharts(filtered);
}

function getFilteredTrades() {
  const from = document.getElementById("dateFrom").value;
  const to = document.getElementById("dateTo").value;
  const symbol = document.getElementById("symbolFilter").value;
  const strategy = document.getElementById("strategyFilter").value;
  const direction = document.getElementById("directionFilter").value;
  const result = document.getElementById("resultFilter").value;

  return trades.filter((trade) => {
    if (from && trade.date && trade.date < from) return false;
    if (to && trade.date && trade.date > to) return false;
    if (symbol !== "all" && trade.symbol !== symbol) return false;
    if (strategy !== "all" && trade.strategy !== strategy) return false;
    if (direction !== "all" && trade.direction !== direction) return false;
    if (result === "win" && trade.netPnl <= 0) return false;
    if (result === "loss" && trade.netPnl >= 0) return false;
    if (result === "flat" && trade.netPnl !== 0) return false;
    return true;
  });
}

function renderKpis(data) {
  const netPnl = sum(data, "netPnl");
  const wins = data.filter((trade) => trade.netPnl > 0);
  const losses = data.filter((trade) => trade.netPnl < 0);
  const grossWin = sum(wins, "netPnl");
  const grossLoss = Math.abs(sum(losses, "netPnl"));
  const avgR = data.length ? sum(data, "rMultiple") / data.length : 0;

  document.getElementById("netPnl").textContent = euro.format(netPnl);
  document.getElementById("netPnl").className = netPnl >= 0 ? "positive" : "negative";
  document.getElementById("winRate").textContent = data.length ? `${number.format((wins.length / data.length) * 100)}%` : "0%";
  document.getElementById("profitFactor").textContent = grossLoss > 0 ? number.format(grossWin / grossLoss) : grossWin > 0 ? "∞" : "0,00";
  document.getElementById("maxDrawdown").textContent = euro.format(maxDrawdown(data));
  document.getElementById("tradeCount").textContent = String(data.length);
  document.getElementById("avgR").textContent = `${number.format(avgR)}R`;
}

function renderTable(data) {
  const body = document.getElementById("tradeTable");
  const rows = data.slice(-100).reverse();
  document.getElementById("shownRows").textContent = `${data.length} Zeilen`;
  body.innerHTML = rows
    .map(
      (trade) => `<tr>
        <td>${escapeHtml(formatDate(trade.date))}</td>
        <td>${escapeHtml(trade.symbol)}</td>
        <td>${escapeHtml(trade.strategy)}</td>
        <td>${escapeHtml(trade.direction)}</td>
        <td>${euro.format(trade.risk)}</td>
        <td class="${trade.netPnl >= 0 ? "positive" : "negative"}">${euro.format(trade.netPnl)}</td>
        <td>${number.format(trade.rMultiple)}R</td>
      </tr>`,
    )
    .join("");
}

function renderCharts(data) {
  const sorted = [...data].sort((a, b) => a.date.localeCompare(b.date));
  let equity = 0;
  const equityLabels = [];
  const equityValues = [];
  sorted.forEach((trade, index) => {
    equity += trade.netPnl;
    equityLabels.push(trade.date ? formatDate(trade.date) : `Trade ${index + 1}`);
    equityValues.push(Number(equity.toFixed(2)));
  });

  upsertChart("equityChart", "line", equityLabels, [{ label: "Equity", data: equityValues, borderColor: "#157f55", backgroundColor: "rgba(21, 127, 85, 0.12)", fill: true, tension: 0.25 }]);
  upsertChart("strategyChart", "bar", ...groupChart(data, "strategy", "#2b6cb0"));
  upsertChart("symbolChart", "bar", ...groupChart(data, "symbol", "#b7791f"));
  upsertChart("directionChart", "doughnut", ...groupChart(data, "direction", ["#157f55", "#c3403d", "#66706a"]));
  upsertChart("monthChart", "bar", ...monthChart(data));
}

function groupChart(data, key, color) {
  const grouped = groupSum(data, key);
  return [Object.keys(grouped), [{ label: "P&L Netto", data: Object.values(grouped), backgroundColor: color }]];
}

function monthChart(data) {
  const grouped = {};
  data.forEach((trade) => {
    const key = trade.date ? trade.date.slice(0, 7) : "Ohne Datum";
    grouped[key] = (grouped[key] || 0) + trade.netPnl;
  });
  const labels = Object.keys(grouped).sort();
  return [labels, [{ label: "P&L Netto", data: labels.map((label) => grouped[label]), backgroundColor: "#157f55" }]];
}

function upsertChart(id, type, labels, datasets) {
  const ctx = document.getElementById(id);
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(ctx, {
    type,
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: type === "doughnut" },
        tooltip: {
          callbacks: {
            label: (context) => `${context.dataset.label || context.label}: ${euro.format(context.raw || 0)}`,
          },
        },
      },
      scales: type === "doughnut" ? {} : { y: { ticks: { callback: (value) => euro.format(value) } } },
    },
  });
}

function groupSum(data, key) {
  return data.reduce((acc, trade) => {
    const label = trade[key] || "Unbekannt";
    acc[label] = (acc[label] || 0) + trade.netPnl;
    return acc;
  }, {});
}

function maxDrawdown(data) {
  let equity = 0;
  let peak = 0;
  let drawdown = 0;
  [...data]
    .sort((a, b) => a.date.localeCompare(b.date))
    .forEach((trade) => {
      equity += trade.netPnl;
      peak = Math.max(peak, equity);
      drawdown = Math.min(drawdown, equity - peak);
    });
  return drawdown;
}

function sum(data, key) {
  return data.reduce((total, trade) => total + (trade[key] || 0), 0);
}

function formatDate(value) {
  if (!value) return "";
  return new Intl.DateTimeFormat("de-DE").format(new Date(`${value}T12:00:00`));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function loadDemo() {
  const strategies = ["Breakout", "Pullback", "Reversal", "Trendfolge"];
  const symbols = ["NQ", "ES", "EURUSD", "BTCUSD", "AAPL"];
  const directions = ["Long", "Short"];
  trades = Array.from({ length: 80 }, (_, index) => {
    const date = new Date(2026, 0, 2 + index * 2);
    const risk = [100, 150, 200, 250][index % 4];
    const r = Math.sin(index * 1.7) > -0.25 ? 0.35 + Math.random() * 2.2 : -0.4 - Math.random() * 1.3;
    const netPnl = Math.round(risk * r - (4 + Math.random() * 9));
    return {
      id: index + 1,
      date: date.toISOString().slice(0, 10),
      symbol: symbols[index % symbols.length],
      market: index % 3 === 0 ? "Futures" : "Aktien/FX",
      direction: directions[index % 2],
      strategy: strategies[index % strategies.length],
      entry: 0,
      exit: 0,
      quantity: 1,
      risk,
      pnl: netPnl,
      fees: 6,
      netPnl,
      rMultiple: netPnl / risk,
    };
  });
  document.getElementById("dataStatus").textContent = "Demo-Daten geladen";
  setupFilters();
  render();
}

function clearData() {
  trades = [];
  document.getElementById("excelFile").value = "";
  document.getElementById("dataStatus").textContent = "Keine Datei geladen";
  resetFilters();
  render();
}
