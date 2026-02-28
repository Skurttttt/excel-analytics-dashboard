/**
 * Frontend logic:
 * - Upload Excel file to backend (/api/upload)
 * - Render KPIs + charts + table
 * - Export as PDF (one long page) INCLUDING header
 * - Client name input (centered) autosaved + used in PDF filename
 *
 * ✅ PDF FIX:
 * Inputs sometimes render faint/blank in html2canvas.
 * Option 1: During export, replace input with a high-contrast "print label" div.
 */

// DOM Elements
const excelFile = document.getElementById("excelFile");
const uploadBtn = document.getElementById("uploadBtn");
const statusEl = document.getElementById("status");
const kpiGrid = document.getElementById("kpiGrid");
const tableBody = document.getElementById("tableBody");
const sheetLabel = document.getElementById("sheetLabel");

const recommendationBox = document.getElementById("recommendationBox");
const targetCacInput = document.getElementById("targetCacInput");

const exportPdfBtn = document.getElementById("exportPdfBtn");
const topHeader = document.getElementById("topHeader");
const pdfRoot = document.getElementById("pdfRoot");

const notesBox = document.getElementById("notesBox");
const clientNameInput = document.getElementById("clientNameInput");

// Keep latest uploaded payload for live updates
window.latestPayload = null;

let charts = {
  spent: null,
  revenue: null,
  messages: null,
  cpm: null,
  pie: null,
};

// -------------------------
// Helpers
// -------------------------
function formatPeso(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return "₱" + n.toLocaleString("en-PH", { maximumFractionDigits: 2 });
}
function formatNumber(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return n.toLocaleString("en-PH", { maximumFractionDigits: 2 });
}
function formatROAS(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "—";
  return n.toLocaleString("en-PH", { maximumFractionDigits: 2 });
}

function setStatus(msg, isError = false) {
  if (!statusEl) return;
  statusEl.textContent = msg;
  statusEl.className =
    "mt-4 text-sm " + (isError ? "text-rose-300" : "text-slate-400");
}

function destroyCharts() {
  Object.values(charts).forEach((c) => c && c.destroy());
  charts = { spent: null, revenue: null, messages: null, cpm: null, pie: null };
}

function createKpiCard({ title, value, sub }) {
  const div = document.createElement("div");
  div.className =
    "rounded-2xl border border-white/10 bg-slate-900/30 p-4 shadow-sm hover:bg-slate-900/40 transition";
  div.innerHTML = `
    <div class="text-xs text-slate-400">${title}</div>
    <div class="mt-2 text-2xl font-semibold">${value}</div>
    <div class="mt-1 text-xs text-slate-400">${sub || ""}</div>
  `;
  return div;
}

function getSeriesKey(series, want) {
  const keys = Object.keys(series || {});
  const w = want.toLowerCase();
  return (
    keys.find((k) => k.toLowerCase() === w) ||
    keys.find((k) => k.toLowerCase().includes(w)) ||
    null
  );
}

/**
 * ✅ IMPORTANT:
 * Do NOT do .map(Number) because Number(null) === 0.
 * We only accept real numbers that are finite and > 0.
 */
function avgFinite(arr) {
  const nums = (arr || []).filter(
    (v) => typeof v === "number" && Number.isFinite(v) && v > 0
  );
  if (!nums.length) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

// ✅ Always treat backend as source of truth for CAC (but compute fallback)
function computeCACLikeExcel(payload) {
  if (!payload || payload.mode !== "overview-style") return null;

  const backendCAC = payload?.kpis?.kpis?.cac;
  if (Number.isFinite(backendCAC)) return backendCAC;

  const s = payload.series || {};
  const keyCAC = getSeriesKey(s, "CAC");
  if (keyCAC && Array.isArray(s[keyCAC])) {
    const avg = avgFinite(s[keyCAC]);
    if (Number.isFinite(avg)) return avg;
  }

  const spent = Number(payload?.kpis?.totals?.spent);
  const cust = Number(payload?.kpis?.totals?.customers);
  if (Number.isFinite(spent) && Number.isFinite(cust) && cust !== 0) return spent / cust;

  return null;
}

function computeCostPerMessageSeries(spentArr = [], msgArr = []) {
  const len = Math.max(spentArr.length, msgArr.length);
  const out = [];
  for (let i = 0; i < len; i++) {
    const s = Number(spentArr[i]);
    const m = Number(msgArr[i]);
    if (!Number.isFinite(s) || !Number.isFinite(m) || m === 0) out.push(null);
    else out.push(s / m);
  }
  return out;
}

// -------------------------
// Data Preview helpers
// -------------------------
function isIsoDateString(v) {
  return (
    typeof v === "string" &&
    /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(v)
  );
}
function formatMonthYear(v) {
  const d = v instanceof Date ? v : new Date(v);
  if (Number.isNaN(d.getTime())) return null;
  return d.toLocaleString("en-US", { month: "long", year: "numeric" });
}
function formatTableNumber(v) {
  const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, "").trim());
  if (!Number.isFinite(n)) return null;
  return n.toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// =========================
// ✅ PDF EXPORT (ONE LONG PAGE) + CLIENT LABEL FIX (OPTION 1)
// =========================

let _printLabelEl = null;
let _savedInputStyle = null;

function mountExportPrintLabel() {
  if (!clientNameInput) return;

  const name = (clientNameInput.value || "").trim() || "Name/Company";

  // Find the wrapper that contains the input
  const wrapper = clientNameInput.parentElement; // <div class="w-full md:max-w-xs">
  if (!wrapper) return;

  // Save input styles so we can restore
  _savedInputStyle = {
    visibility: clientNameInput.style.visibility,
    opacity: clientNameInput.style.opacity,
  };

  // Hide the real input (but keep layout)
  clientNameInput.style.visibility = "hidden";
  clientNameInput.style.opacity = "0";

  // Create a print-only label on top
  _printLabelEl = document.createElement("div");
  _printLabelEl.id = "clientNamePrintLabel";
  _printLabelEl.textContent = name;

  // Strong high-contrast style for PDF
  _printLabelEl.style.position = "absolute";
  _printLabelEl.style.left = "0";
  _printLabelEl.style.top = "0";
  _printLabelEl.style.width = "100%";
  _printLabelEl.style.height = "100%";
  _printLabelEl.style.display = "flex";
  _printLabelEl.style.alignItems = "center";
  _printLabelEl.style.justifyContent = "center";
  _printLabelEl.style.textAlign = "center";
  _printLabelEl.style.padding = "10px 12px";
  _printLabelEl.style.borderRadius = "12px";
  _printLabelEl.style.background = "#0b1220"; // solid (no blur)
  _printLabelEl.style.border = "1px solid rgba(255,255,255,0.18)";
  _printLabelEl.style.color = "#ffffff";
  _printLabelEl.style.fontWeight = "700";
  _printLabelEl.style.fontSize = "14px";
  _printLabelEl.style.letterSpacing = "0.2px";
  _printLabelEl.style.lineHeight = "1.1";
  _printLabelEl.style.boxShadow = "0 8px 30px rgba(0,0,0,0.35)";

  // Ensure wrapper can position absolute child
  const wrapperComputed = window.getComputedStyle(wrapper);
  if (wrapperComputed.position === "static") {
    wrapper.dataset._pos = wrapper.style.position;
    wrapper.style.position = "relative";
  }

  wrapper.appendChild(_printLabelEl);
}

function unmountExportPrintLabel() {
  if (clientNameInput && _savedInputStyle) {
    clientNameInput.style.visibility = _savedInputStyle.visibility || "";
    clientNameInput.style.opacity = _savedInputStyle.opacity || "";
  }
  _savedInputStyle = null;

  if (_printLabelEl && _printLabelEl.parentElement) {
    _printLabelEl.parentElement.removeChild(_printLabelEl);
  }
  _printLabelEl = null;

  // restore wrapper position if we changed it
  if (clientNameInput?.parentElement?.dataset?._pos !== undefined) {
    clientNameInput.parentElement.style.position = clientNameInput.parentElement.dataset._pos || "";
    delete clientNameInput.parentElement.dataset._pos;
  }
}

function setExportMode(enable) {
  // Expand scroll containers for capture
  const scrollEls = document.querySelectorAll(".overflow-auto");
  scrollEls.forEach((el) => {
    if (enable) {
      el.dataset._overflow = el.style.overflow;
      el.dataset._maxHeight = el.style.maxHeight;
      el.dataset._height = el.style.height;

      el.style.overflow = "visible";
      el.style.maxHeight = "none";
      el.style.height = "auto";
    } else {
      el.style.overflow = el.dataset._overflow || "";
      el.style.maxHeight = el.dataset._maxHeight || "";
      el.style.height = el.dataset._height || "";

      delete el.dataset._overflow;
      delete el.dataset._maxHeight;
      delete el.dataset._height;
    }
  });

  // Keep header visible but remove sticky for clean capture
  if (topHeader) {
    if (enable) {
      topHeader.dataset._position = topHeader.style.position;
      topHeader.dataset._top = topHeader.style.top;
      topHeader.dataset._z = topHeader.style.zIndex;
      topHeader.dataset._bg = topHeader.style.background;

      topHeader.style.position = "relative";
      topHeader.style.top = "0";
      topHeader.style.zIndex = "1";
      // Force solid bg (backdrop blur can wash text in canvas)
      topHeader.style.background = "#020617";
    } else {
      topHeader.style.position = topHeader.dataset._position || "";
      topHeader.style.top = topHeader.dataset._top || "";
      topHeader.style.zIndex = topHeader.dataset._z || "";
      topHeader.style.background = topHeader.dataset._bg || "";

      delete topHeader.dataset._position;
      delete topHeader.dataset._top;
      delete topHeader.dataset._z;
      delete topHeader.dataset._bg;
    }
  }

  // ✅ Option 1 print label
  if (enable) mountExportPrintLabel();
  else unmountExportPrintLabel();
}

async function exportDashboardOnePagePDF() {
  if (!pdfRoot) return;

  try {
    setStatus("Exporting PDF (one long page)...");

    setExportMode(true);

    // Stabilize layout (charts)
    window.dispatchEvent(new Event("resize"));
    await new Promise((r) => setTimeout(r, 350));

    const canvas = await html2canvas(pdfRoot, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#020617",
      scrollX: 0,
      scrollY: -window.scrollY,
      logging: false,
      allowTaint: false,
    });

    const imgData = canvas.toDataURL("image/png");
    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "px",
      format: [canvas.width, canvas.height],
      compress: true,
    });

    pdf.addImage(imgData, "PNG", 0, 0, canvas.width, canvas.height, undefined, "FAST");

    const clientName = (clientNameInput?.value || "").trim();
    const safeClient = clientName
      ? clientName.replace(/[^\w\- ]+/g, "").replace(/\s+/g, "-")
      : "Client";

    const fileName = `Digital-Homie-Analytics-${safeClient}-${new Date()
      .toISOString()
      .slice(0, 10)}.pdf`;

    pdf.save(fileName);

    setStatus("PDF exported successfully ✅");
  } catch (err) {
    console.error(err);
    setStatus("Export failed: " + err.message, true);
  } finally {
    setExportMode(false);
  }
}

// -------------------------
// KPI render
// -------------------------
function renderKPIs(payload) {
  kpiGrid.innerHTML = "";

  if (payload.mode !== "overview-style") {
    kpiGrid.appendChild(
      createKpiCard({ title: "KPIs", value: "—", sub: "Not in overview-style mode" })
    );
    return;
  }

  const { totals, averagesPerMonth, kpis } = payload.kpis;
  const excelCAC = computeCACLikeExcel(payload);

  kpiGrid.appendChild(
    createKpiCard({
      title: "Ad Spent",
      value: formatPeso(totals.spent),
      sub: `Avg/mo ${formatPeso(averagesPerMonth.spent)}`,
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "Messages",
      value: formatNumber(totals.messages),
      sub: `Avg/mo ${formatNumber(averagesPerMonth.messages)}`,
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "Revenue",
      value: formatPeso(totals.revenue),
      sub: `Avg/mo ${formatPeso(averagesPerMonth.revenue)}`,
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "Cost/Message",
      value: formatPeso(kpis.costPerMessage),
      sub: "(Spent ÷ Messages)",
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "ROAS",
      value: formatROAS(kpis.roas),
      sub: "(Revenue ÷ Spent)",
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "CAC",
      value: formatPeso(excelCAC),
      sub: "(Avg of CAC row values)",
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "Customers",
      value: formatNumber(totals.customers),
      sub: `Avg/mo ${formatNumber(averagesPerMonth.customers)}`,
    })
  );
}

// -------------------------
// Charts + recommendations (your existing logic)
// -------------------------
function renderCharts(payload) {
  destroyCharts();
  if (payload.mode !== "overview-style") return;

  const { labels, series } = payload;

  const keySpent =
    getSeriesKey(series, "Total Ad Spent") ||
    getSeriesKey(series, "Amount spent") ||
    getSeriesKey(series, "Ad Spent");

  const keyRevenue =
    getSeriesKey(series, "Total Revenue") || getSeriesKey(series, "Revenue");

  const keyMessages =
    getSeriesKey(series, "No. of Messages") || getSeriesKey(series, "Messages");

  const keyCTR =
    getSeriesKey(series, "CTR") ||
    getSeriesKey(series, "Link CTR") ||
    getSeriesKey(series, "CTR (Link)");

  const keyCPM =
    getSeriesKey(series, "Cost Per Message") ||
    getSeriesKey(series, "Cost/Message") ||
    getSeriesKey(series, "Cost per message");

  const spentData = keySpent ? series[keySpent] : [];
  const revenueData = keyRevenue ? series[keyRevenue] : [];
  const msgData = keyMessages ? series[keyMessages] : [];
  const ctrData = keyCTR ? series[keyCTR] : [];

  const cpmFromSheet = keyCPM ? series[keyCPM] : null;
  const computedCPM = computeCostPerMessageSeries(spentData, msgData);

  Chart.defaults.color = "rgba(226, 232, 240, 0.85)";
  Chart.defaults.borderColor = "rgba(255,255,255,0.10)";

  charts.spent = new Chart(document.getElementById("chartSpent"), {
    type: "line",
    data: { labels, datasets: [{ label: "Ad Spent", data: spentData, tension: 0.35, borderWidth: 3, pointRadius: 3, borderColor: "#60A5FA", backgroundColor: "rgba(96,165,250,0.20)", fill: true }] },
    options: { responsive: true, plugins: { legend: { display: false } } },
  });

  charts.revenue = new Chart(document.getElementById("chartRevenue"), {
    type: "line",
    data: { labels, datasets: [{ label: "Revenue", data: revenueData, tension: 0.35, borderWidth: 3, pointRadius: 3, borderColor: "#A78BFA", backgroundColor: "rgba(167,139,250,0.20)", fill: true }] },
    options: { responsive: true, plugins: { legend: { display: false } } },
  });

  charts.messages = new Chart(document.getElementById("chartMessages"), {
    type: "bar",
    data: { labels, datasets: [{ label: "Messages", data: msgData, borderWidth: 1, backgroundColor: "rgba(52,211,153,0.45)", borderColor: "rgba(52,211,153,0.95)" }] },
    options: { responsive: true, plugins: { legend: { display: false } } },
  });

  charts.cpm = new Chart(document.getElementById("chartCostPerMessage"), {
    type: "line",
    data: { labels, datasets: [{ label: "Cost / Message", data: computedCPM, tension: 0.35, borderWidth: 3, pointRadius: 3, borderColor: "#FBBF24", backgroundColor: "rgba(251,191,36,0.18)", fill: true }] },
    options: { responsive: true, plugins: { legend: { display: false } } },
  });

  const totalSpent = payload.kpis?.totals?.spent ?? 0;
  const totalRevenue = payload.kpis?.totals?.revenue ?? 0;

  charts.pie = new Chart(document.getElementById("chartPie"), {
    type: "pie",
    data: { labels: ["Ad Spent", "Revenue"], datasets: [{ data: [totalSpent, totalRevenue], borderWidth: 1, backgroundColor: ["rgba(96,165,250,0.55)", "rgba(167,139,250,0.55)"], borderColor: ["rgba(96,165,250,1)", "rgba(167,139,250,1)"] }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "top" } } },
  });

  renderRecommendations(payload, { ctrData, cpmFromSheet, computedCPM });
}

function renderRecommendations(payload, { ctrData = [], cpmFromSheet = null, computedCPM = [] } = {}) {
  if (!recommendationBox) return;

  const roas = payload.kpis?.kpis?.roas;
  const cac = computeCACLikeExcel(payload);
  const costPerMessage = payload.kpis?.kpis?.costPerMessage;

  const targetCAC = Number(targetCacInput?.value);

  const ctrNums = (ctrData || []).map(Number).filter(Number.isFinite);
  const avgCTR = ctrNums.length ? ctrNums.reduce((a, b) => a + b, 0) / ctrNums.length : null;

  const recs = [];
  const addRec = (level, title, why, actions) => {
    const icon = level === "high" ? "🔴" : level === "med" ? "🟡" : "🟢";
    recs.push({ icon, title, why, actions });
  };

  if (Number.isFinite(roas) && roas < 3) {
    addRec("high", "Campaign not profitable", `ROAS is ${roas.toFixed(2)} (target 3.00+)`, [
      "Test 3 new hooks (UGC / problem-solution / proof)",
      "Tighten targeting (exclude low-quality audiences)",
      "Improve offer + landing page conversion",
    ]);
  }

  if (Number.isFinite(avgCTR) && avgCTR < 1) {
    addRec("med", "Low CTR (weak hook)", `Average CTR is ${avgCTR.toFixed(2)}% (target 1%+)`, [
      "Rewrite first 2 lines (strong hook + pain)",
      "Try new thumbnails / opening frame",
      "Use benefit-led headline + clear CTA",
    ]);
  }

  if (Number.isFinite(cac) && Number.isFinite(targetCAC) && cac > targetCAC) {
    addRec("high", "CAC above target", `CAC is ₱${cac.toLocaleString("en-PH", { maximumFractionDigits: 2 })} vs target ₱${targetCAC.toLocaleString("en-PH")}`, [
      "Refresh creatives (new angles every 7–10 days)",
      "Improve conversion rate (offer + follow-up speed)",
      "Retarget warm users (video viewers, engagers, past messages)",
    ]);
  }

  let cpmArr = null;
  if (Array.isArray(cpmFromSheet) && cpmFromSheet.length >= 2) cpmArr = cpmFromSheet.map(Number);
  else if (Array.isArray(computedCPM) && computedCPM.length >= 2) cpmArr = computedCPM.map(Number);

  if (cpmArr) {
    const nums = cpmArr.filter((x) => Number.isFinite(x));
    if (nums.length >= 2) {
      const last = nums[nums.length - 1];
      const prev = nums[nums.length - 2];
      if (Number.isFinite(last) && Number.isFinite(prev) && last > prev * 1.2) {
        addRec("med", "Cost per message rising", `Last ₱${last.toFixed(2)} vs prev ₱${prev.toFixed(2)} (+20%+)`, [
          "Refresh creatives (new hook + new first frame)",
          "Check audience fatigue (frequency/exclusions)",
          "Improve response speed (slow replies reduce conversion)",
        ]);
      }
    }
  }

  if (recs.length === 0) {
    addRec("low", "Healthy signals", "No major red flags detected.", [
      "Keep testing creatives weekly",
      "Monitor CAC + ROAS trends monthly",
    ]);
  }

  recommendationBox.innerHTML = `
    <div class="flex flex-wrap gap-2">
      <span class="px-2 py-1 rounded-lg bg-white/5 border border-white/10 text-xs">
        ROAS: <b>${Number.isFinite(roas) ? roas.toFixed(2) : "—"}</b>
      </span>
      <span class="px-2 py-1 rounded-lg bg-white/5 border border-white/10 text-xs">
        CAC: <b>${Number.isFinite(cac) ? "₱" + cac.toLocaleString("en-PH", { maximumFractionDigits: 2 }) : "—"}</b>
      </span>
      <span class="px-2 py-1 rounded-lg bg-white/5 border border-white/10 text-xs">
        Cost/Msg: <b>${Number.isFinite(costPerMessage) ? "₱" + Number(costPerMessage).toFixed(2) : "—"}</b>
      </span>
      ${
        Number.isFinite(targetCAC)
          ? `<span class="px-2 py-1 rounded-lg bg-indigo-500/10 border border-indigo-400/20 text-xs">
               Target CAC: <b>₱${targetCAC.toLocaleString("en-PH")}</b>
             </span>`
          : ""
      }
    </div>
  `;
}

// -------------------------
// Table preview (same as yours)
// -------------------------
function renderTablePreview(payload) {
  tableBody.innerHTML = "";
  const rows = payload.tablePreview?.rows || [];
  if (!rows.length) return;

  rows.forEach((r, idx) => {
    const tr = document.createElement("tr");
    tr.className = idx === 0 ? "bg-white/10 text-white font-semibold" : "hover:bg-white/5 transition";

    (r || []).slice(0, 25).forEach((cell, colIdx) => {
      const td = document.createElement("td");
      td.className = "px-3 py-2 text-xs text-slate-200 border border-white/10 whitespace-nowrap";
      if (colIdx === 0) td.className += " font-medium text-white bg-white/5";

      if (isIsoDateString(cell) || cell instanceof Date) {
        td.textContent = formatMonthYear(cell) ?? "";
        tr.appendChild(td);
        return;
      }

      const num = formatTableNumber(cell);
      if (num !== null) {
        td.textContent = num;
        td.className += " text-right";
        tr.appendChild(td);
        return;
      }

      td.textContent = cell === null || cell === undefined ? "" : String(cell);
      tr.appendChild(td);
    });

    tableBody.appendChild(tr);
  });
}

// -------------------------
// Upload + Render
// -------------------------
async function uploadAndRender() {
  const file = excelFile?.files?.[0];
  if (!file) return setStatus("Please choose an .xlsx file first.", true);

  uploadBtn.disabled = true;
  setStatus("Uploading and processing...");

  const formData = new FormData();
  formData.append("excel", file);

  try {
    const res = await fetch("/api/upload", { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Upload failed.");

    window.latestPayload = data;
    if (sheetLabel) sheetLabel.textContent = `Sheet: ${data.sheetName} | Mode: ${data.mode}`;

    renderKPIs(data);
    renderCharts(data);
    renderTablePreview(data);

    setStatus("Dashboard generated successfully ✅");
  } catch (err) {
    console.error(err);
    setStatus(err.message, true);
  } finally {
    uploadBtn.disabled = false;
  }
}

// -------------------------
// Events
// -------------------------
uploadBtn?.addEventListener("click", uploadAndRender);

targetCacInput?.addEventListener("input", () => {
  if (window.latestPayload) renderCharts(window.latestPayload);
});

// Notes autosave
if (notesBox) {
  const saved = localStorage.getItem("excel_dashboard_notes");
  if (saved) notesBox.value = saved;
  notesBox.addEventListener("input", () => {
    localStorage.setItem("excel_dashboard_notes", notesBox.value);
  });
}

// Client name autosave
if (clientNameInput) {
  const savedClient = localStorage.getItem("excel_dashboard_client_name");
  if (savedClient) clientNameInput.value = savedClient;
  clientNameInput.addEventListener("input", () => {
    localStorage.setItem("excel_dashboard_client_name", clientNameInput.value);
  });
}

// Export PDF
exportPdfBtn?.addEventListener("click", exportDashboardOnePagePDF);