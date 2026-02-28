const express = require("express");
const path = require("path");
const fs = require("fs");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

// Serve frontend (public/)
app.use(express.static(path.join(__dirname, "public")));

// Ensure uploads folder exists
const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// Multer storage
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) =>
    cb(null, `${Date.now()}_${file.originalname.replace(/\s+/g, "_")}`),
});

function fileFilter(req, file, cb) {
  const ok = file.originalname.toLowerCase().endsWith(".xlsx");
  if (!ok) return cb(new Error("Only .xlsx files are allowed."), false);
  cb(null, true);
}

const upload = multer({ storage, fileFilter });

// ✅ FIXED toNumber: empty string does NOT become 0
function toNumber(v) {
  if (v === null || v === undefined) return null;

  if (typeof v === "string") {
    const raw = v.trim();
    if (!raw || raw === "-" || raw === "—") return null;
  }

  if (typeof v === "number") return Number.isFinite(v) ? v : null;

  const cleaned = String(v).replace(/[₱,$\s]/g, "").replace(/,/g, "").trim();
  if (!cleaned) return null;

  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function safeDivide(a, b) {
  if (!Number.isFinite(a) || !Number.isFinite(b) || b === 0) return null;
  return a / b;
}

function sumFinite(arr) {
  return (arr || []).filter(Number.isFinite).reduce((a, b) => a + b, 0);
}
function countFinite(arr) {
  return (arr || []).filter(Number.isFinite).length;
}
function avgFinite(arr) {
  // ✅ ignore zeros too (blank months sometimes end up as 0 in messy files)
  const nums = (arr || []).filter((n) => Number.isFinite(n) && n > 0);
  if (!nums.length) return null;
  return nums.reduce((a, b) => a + b, 0) / nums.length;
}

function parseOverviewStyle(aoa) {
  let headerRowIndex = -1;

  for (let r = 0; r < Math.min(aoa.length, 60); r++) {
    const row = aoa[r] || [];
    const dateLikeCount = row.slice(1).filter((cell) => {
      if (cell instanceof Date) return true;
      if (typeof cell === "number") return cell > 20000 && cell < 60000;
      if (typeof cell === "string") return /\b(20\d{2}|19\d{2})\b/.test(cell);
      return false;
    }).length;

    if (dateLikeCount >= 3) {
      headerRowIndex = r;
      break;
    }
  }

  if (headerRowIndex === -1) return null;

  const headerRow = aoa[headerRowIndex] || [];
  const labels = headerRow.slice(1).map((cell) => {
    if (cell instanceof Date) {
      return cell.toLocaleString("en-US", { month: "short", year: "2-digit" });
    }
    if (typeof cell === "number") {
      const d = XLSX.SSF.parse_date_code(cell);
      if (d) {
        return new Date(d.y, d.m - 1, d.d).toLocaleString("en-US", {
          month: "short",
          year: "2-digit",
        });
      }
    }
    return String(cell ?? "").trim() || "—";
  });

  const series = {};
  for (let r = headerRowIndex + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const metricName = String(row[0] ?? "").trim();
    if (!metricName) continue;

    const values = row.slice(1).map(toNumber);
    if (!values.some((v) => Number.isFinite(v))) continue;

    series[metricName] = values;
  }

  return { labels, series };
}

function computeKPIsFromSeries(labels, series) {
  const findKey = (name) => {
    const target = name.toLowerCase();
    return (
      Object.keys(series).find((k) => k.toLowerCase() === target) ||
      Object.keys(series).find((k) => k.toLowerCase().includes(target))
    );
  };

  const keySpent =
    findKey("Total Ad Spent") || findKey("Amount spent") || findKey("Ad Spent");
  const keyMsgs = findKey("No. of Messages") || findKey("Messages");
  const keyRevenue = findKey("Total Revenue") || findKey("Revenue");
  const keyCustomers = findKey("No. of Customers") || findKey("Customers");
  const keyCAC = findKey("CAC"); // ✅ CAC row

  const spentArr = keySpent ? series[keySpent] : [];
  const msgArr = keyMsgs ? series[keyMsgs] : [];
  const revArr = keyRevenue ? series[keyRevenue] : [];
  const custArr = keyCustomers ? series[keyCustomers] : [];
  const cacArr = keyCAC ? series[keyCAC] : [];

  const spent = keySpent ? sumFinite(spentArr) : null;
  const messages = keyMsgs ? sumFinite(msgArr) : null;
  const revenue = keyRevenue ? sumFinite(revArr) : null;
  const customers = keyCustomers ? sumFinite(custArr) : null;

  // per-metric month counts
  const spentMonths = Math.max(countFinite(spentArr), 1);
  const msgMonths = Math.max(countFinite(msgArr), 1);
  const revMonths = Math.max(countFinite(revArr), 1);
  const custMonths = Math.max(countFinite(custArr), 1);

  const costPerMessage = safeDivide(spent, messages);
  const roas = safeDivide(revenue, spent);

  // ✅ FORCE CAC = Excel average of CAC row values (sum ÷ count of months with CAC)
  const cacExcelStyle = avgFinite(cacArr);
  const cacBlended = safeDivide(spent, customers);

  return {
    totals: { spent, messages, revenue, customers },
    averagesPerMonth: {
      spent: spent !== null ? spent / spentMonths : null,
      messages: messages !== null ? messages / msgMonths : null,
      revenue: revenue !== null ? revenue / revMonths : null,
      customers: customers !== null ? customers / custMonths : null,
    },
    kpis: {
      costPerMessage,
      roas,
      cac: Number.isFinite(cacExcelStyle) ? cacExcelStyle : cacBlended,
    },
  };
}

// Upload endpoint
app.post("/api/upload", upload.single("excel"), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded." });

    const workbook = XLSX.readFile(req.file.path, { cellDates: true });
    const sheetName =
      workbook.SheetNames.find((n) => n.toLowerCase() === "overview") ||
      workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });
    const parsed = parseOverviewStyle(aoa);

    const tablePreview = { sheetName, rows: aoa.slice(0, 50) };

    // delete uploaded file after parsing
    fs.unlink(req.file.path, () => {});

    if (!parsed) {
      return res.status(422).json({
        error: "Could not detect Overview-style format. Please use the provided sample format.",
        tablePreview,
      });
    }

    const { labels, series } = parsed;
    const kpis = computeKPIsFromSeries(labels, series);

    res.json({ mode: "overview-style", sheetName, labels, series, kpis, tablePreview });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Server error." });
  }
});

app.listen(PORT, () => console.log(`✅ Server running on http://localhost:${PORT}`));