/**
 * Frontend logic:
 * - Upload Excel file to backend (/api/upload)
 * - Receive JSON (KPIs + labels + series)
 * - Render KPI cards
 * - Render charts using Chart.js
 * - Render table preview (with grid lines)
 * - Export as PDF (screenshot-style)
 * - Notes box under table (auto-save)
 */

const excelFile = document.getElementById("excelFile");
const uploadBtn = document.getElementById("uploadBtn");
const statusEl = document.getElementById("status");
const kpiGrid = document.getElementById("kpiGrid");
const tableBody = document.getElementById("tableBody");
const sheetLabel = document.getElementById("sheetLabel");

const recommendationBox = document.getElementById("recommendationBox");
const targetCacInput = document.getElementById("targetCacInput");

const exportPdfBtn = document.getElementById("exportPdfBtn");
const dashboardRoot = document.getElementById("dashboardRoot");

const notesBox = document.getElementById("notesBox");

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
// Basic number format helpers
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

// -------------------------
// Data Preview formatting helpers
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
  return d.toLocaleString("en-US", { month: "long", year: "numeric" }); // "June 2025"
}

function formatTableNumber(v) {
  const n =
    typeof v === "number" ? v : Number(String(v).replace(/,/g, "").trim());
  if (!Number.isFinite(n)) return null;
  return n.toLocaleString("en-PH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

// -------------------------
function setStatus(msg, isError = false) {
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

function renderKPIs(payload) {
  kpiGrid.innerHTML = "";

  if (payload.mode !== "overview-style") {
    const stats = payload.tableParsed?.stats || {};
    const entries = Object.entries(stats);

    if (entries.length === 0) {
      kpiGrid.appendChild(
        createKpiCard({
          title: "KPIs",
          value: "—",
          sub: "No KPI metrics detected",
        })
      );
      return;
    }

    entries.slice(0, 12).forEach(([name, s]) => {
      kpiGrid.appendChild(
        createKpiCard({
          title: name,
          value: formatNumber(s.sum),
          sub: `Avg: ${formatNumber(s.avg)} | Count: ${s.count}`,
        })
      );
    });
    return;
  }

  const { totals, averagesPerMonth, kpis } = payload.kpis;

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
      sub: "Computed (Spent ÷ Messages)",
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "ROAS",
      value: formatROAS(kpis.roas),
      sub: "Computed (Revenue ÷ Spent)",
    })
  );

  kpiGrid.appendChild(
    createKpiCard({
      title: "CAC",
      value: formatPeso(kpis.cac),
      sub: "Computed (Spent ÷ Customers)",
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

function getSeriesKey(series, want) {
  const keys = Object.keys(series || {});
  const w = want.toLowerCase();
  return (
    keys.find((k) => k.toLowerCase() === w) ||
    keys.find((k) => k.toLowerCase().includes(w)) ||
    null
  );
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

  // Dark theme defaults
  Chart.defaults.color = "rgba(226, 232, 240, 0.85)";
  Chart.defaults.borderColor = "rgba(255,255,255,0.10)";

  charts.spent = new Chart(document.getElementById("chartSpent"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Ad Spent",
          data: spentData,
          tension: 0.35,
          borderWidth: 3,
          pointRadius: 3,
          borderColor: "#60A5FA",
          backgroundColor: "rgba(96,165,250,0.20)",
          fill: true,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: { y: { ticks: { callback: (v) => "₱" + v } } },
    },
  });

  charts.revenue = new Chart(document.getElementById("chartRevenue"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Revenue",
          data: revenueData,
          tension: 0.35,
          borderWidth: 3,
          pointRadius: 3,
          borderColor: "#A78BFA",
          backgroundColor: "rgba(167,139,250,0.20)",
          fill: true,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: { y: { ticks: { callback: (v) => "₱" + v } } },
    },
  });

  charts.messages = new Chart(document.getElementById("chartMessages"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Messages",
          data: msgData,
          borderWidth: 1,
          backgroundColor: "rgba(52,211,153,0.45)",
          borderColor: "rgba(52,211,153,0.95)",
        },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
    },
  });

  charts.cpm = new Chart(document.getElementById("chartCostPerMessage"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Cost / Message",
          data: computedCPM,
          tension: 0.35,
          borderWidth: 3,
          pointRadius: 3,
          borderColor: "#FBBF24",
          backgroundColor: "rgba(251,191,36,0.18)",
          fill: true,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: { y: { ticks: { callback: (v) => "₱" + v } } },
    },
  });

  const totalSpent = payload.kpis?.totals?.spent ?? 0;
  const totalRevenue = payload.kpis?.totals?.revenue ?? 0;

  charts.pie = new Chart(document.getElementById("chartPie"), {
    type: "pie",
    data: {
      labels: ["Ad Spent", "Revenue"],
      datasets: [
        {
          data: [totalSpent, totalRevenue],
          borderWidth: 1,
          backgroundColor: ["rgba(96,165,250,0.55)", "rgba(167,139,250,0.55)"],
          borderColor: ["rgba(96,165,250,1)", "rgba(167,139,250,1)"],
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: "top" } },
    },
  });

  renderRecommendations(payload, { ctrData, cpmFromSheet, computedCPM });
}

function renderRecommendations(payload, { ctrData = [], cpmFromSheet = null, computedCPM = [] } = {}) {
  if (!recommendationBox) return;

  const roas = payload.kpis?.kpis?.roas;
  const cac = payload.kpis?.kpis?.cac;
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
    addRec("high", "CAC above target", `CAC is ₱${cac.toLocaleString("en-PH")} vs target ₱${targetCAC.toLocaleString("en-PH")}`, [
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

  // ✅ UI-friendly output (chips + cards)
  recommendationBox.innerHTML = `
    <div class="flex flex-wrap gap-2">
      <span class="px-2 py-1 rounded-lg bg-white/5 border border-white/10 text-xs">
        ROAS: <b>${Number.isFinite(roas) ? roas.toFixed(2) : "—"}</b>
      </span>
      <span class="px-2 py-1 rounded-lg bg-white/5 border border-white/10 text-xs">
        CAC: <b>${Number.isFinite(cac) ? "₱" + cac.toLocaleString("en-PH") : "—"}</b>
      </span>
      <span class="px-2 py-1 rounded-lg bg-white/5 border border-white/10 text-xs">
        Cost/Msg: <b>${Number.isFinite(costPerMessage) ? "₱" + Number(costPerMessage).toFixed(2) : "—"}</b>
      </span>
      ${Number.isFinite(targetCAC) ? `
        <span class="px-2 py-1 rounded-lg bg-indigo-500/10 border border-indigo-400/20 text-xs">
          Target CAC: <b>₱${targetCAC.toLocaleString("en-PH")}</b>
        </span>` : ""}
    </div>

    <div class="mt-3 space-y-2">
      ${recs.map((r) => {
        const badge =
          r.icon.includes("🔴") ? "bg-rose-500/15 text-rose-200 border-rose-400/20" :
          r.icon.includes("🟡") ? "bg-amber-500/15 text-amber-200 border-amber-400/20" :
          "bg-emerald-500/15 text-emerald-200 border-emerald-400/20";

        return `
          <div class="rounded-xl border border-white/10 bg-white/5 p-3">
            <div class="flex items-center justify-between gap-2">
              <div class="font-semibold text-sm">${r.title}</div>
              <span class="text-[11px] px-2 py-1 rounded-full border ${badge}">
                ${r.icon.includes("🔴") ? "High" : r.icon.includes("🟡") ? "Medium" : "Good"}
              </span>
            </div>
            <div class="mt-1 text-xs text-slate-300">${r.why}</div>
            <div class="mt-2 text-xs text-slate-200">
              <div class="font-semibold text-slate-300 mb-1">Next actions</div>
              <ul class="list-disc pl-5 space-y-1">
                ${r.actions.map(a => `<li>${a}</li>`).join("")}
              </ul>
            </div>
          </div>
        `;
      }).join("")}
    </div>
  `;
}

function renderTablePreview(payload) {
  tableBody.innerHTML = "";

  const rows = payload.tablePreview?.rows || [];
  if (rows.length === 0) return;

  rows.forEach((r, idx) => {
    const tr = document.createElement("tr");

    if (idx === 0) {
      tr.className = "bg-white/10 text-white font-semibold";
    } else {
      tr.className = "hover:bg-white/5 transition";
    }

    (r || []).slice(0, 25).forEach((cell, colIdx) => {
      const td = document.createElement("td");

      td.className = "px-3 py-2 text-xs text-slate-200 border border-white/10 whitespace-nowrap";

      // First column stronger
      if (colIdx === 0) {
        td.className += " font-medium text-white bg-white/5";
      }

      // ISO date -> "June 2025"
      if (isIsoDateString(cell) || cell instanceof Date) {
        const pretty = formatMonthYear(cell);
        td.textContent = pretty ?? "";
        tr.appendChild(td);
        return;
      }

      // Number -> comma + 2 decimals
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

async function uploadAndRender() {
  const file = excelFile.files[0];
  if (!file) {
    setStatus("Please choose an .xlsx file first.", true);
    return;
  }

  uploadBtn.disabled = true;
  setStatus("Uploading and processing...");

  const formData = new FormData();
  formData.append("excel", file);

  try {
    const res = await fetch("/api/upload", { method: "POST", body: formData });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Upload failed.");

    window.latestPayload = data;

    sheetLabel.textContent = `Sheet: ${data.sheetName} | Mode: ${data.mode}`;

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

uploadBtn.addEventListener("click", uploadAndRender);

// ✅ Live update recommendations when Target CAC changes
if (targetCacInput) {
  targetCacInput.addEventListener("input", () => {
    if (window.latestPayload) renderCharts(window.latestPayload);
  });
}

// ✅ Notes box auto-save (localStorage)
if (notesBox) {
  const saved = localStorage.getItem("excel_dashboard_notes");
  if (saved) notesBox.value = saved;

  notesBox.addEventListener("input", () => {
    localStorage.setItem("excel_dashboard_notes", notesBox.value);
  });
}

// ✅ Export as PDF (screenshot-style)
async function exportAsPDF() {
  if (!dashboardRoot) return;

  try {
    setStatus("Exporting PDF... (please wait)");

    // Temporarily hide the sticky header shadow effect issues? (optional)
    const canvas = await html2canvas(dashboardRoot, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#020617", // slate-950
      scrollX: 0,
      scrollY: -window.scrollY,
    });

    const imgData = canvas.toDataURL("image/png");

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p", "mm", "a4");

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    // Image dimensions in PDF
    const imgProps = pdf.getImageProperties(imgData);
    const imgWidth = pdfWidth;
    const imgHeight = (imgProps.height * imgWidth) / imgProps.width;

    // Multi-page support
    let heightLeft = imgHeight;
    let position = 0;

    pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
    heightLeft -= pdfHeight;

    while (heightLeft > 0) {
      position -= pdfHeight;
      pdf.addPage();
      pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
      heightLeft -= pdfHeight;
    }

    const fileName = `Excel-Dashboard-${new Date().toISOString().slice(0, 10)}.pdf`;
    pdf.save(fileName);

    setStatus("PDF exported successfully ✅");
  } catch (err) {
    console.error(err);
    setStatus("Export failed: " + err.message, true);
  }
}

if (exportPdfBtn) exportPdfBtn.addEventListener("click", exportAsPDF);