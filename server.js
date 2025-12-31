const express = require("express");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

/* ---------------- CORS (SHOPIFY SAFE) ---------------- */
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept"
  );
  res.header(
    "Access-Control-Allow-Methods",
    "GET, POST, OPTIONS"
  );

  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }

  next();
});

/* ---------------- MIDDLEWARE ---------------- */
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

/* ---------------- HEALTH CHECK ---------------- */
app.get("/", (req, res) => {
  res.send("Code Verification Server Running");
});

/* ---------------- LOAD ALL CODES FROM ALL FILES ---------------- */
function loadAllCodes() {
  const folderPath = path.join(__dirname, "codes");
  if (!fs.existsSync(folderPath)) return [];

  const files = fs.readdirSync(folderPath);
  let allCodes = [];

  files.forEach(file => {
    if (!file.endsWith(".xlsx")) return;

    const workbook = XLSX.readFile(path.join(folderPath, file));
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    rows.forEach(row => {
      if (!row.code) return;

      allCodes.push({
        code: String(row.code).trim(),
        used: String(row.used).trim().toUpperCase() === "YES",
        file
      });
    });
  });

  return allCodes;
}

/* ---------------- MARK CODE AS USED + SAVE DETAILS ---------------- */
function markCodeAsUsed(code, file, details) {
  const filePath = path.join(__dirname, "codes", file);
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  const now = new Date();
  const dateUsed = now.toISOString().slice(0, 19).replace("T", " ");

  rows.forEach(row => {
    if (String(row.code).trim() === code) {
      row.used = "YES";
      row.name = details.name || "";
      row.mobile = details.mobile || "";
      row.source = details.purchaseSource || "";
      row.date_used = dateUsed;
    }
  });

  workbook.Sheets[sheetName] = XLSX.utils.json_to_sheet(rows);
  XLSX.writeFile(workbook, filePath);
}

/* ---------------- VERIFY ROUTE ---------------- */
app.post("/verify", (req, res) => {
  try {
    const { code, name, mobile, purchaseSource } = req.body;

    if (!code) {
      return res.json({
        success: false,
        message: "Code is required"
      });
    }

    const cleanCode = String(code).trim();
    const allCodes = loadAllCodes();

    const found = allCodes.find(
      c => c.code === cleanCode && !c.used
    );

    if (!found) {
      return res.json({
        success: false,
        message: "Invalid or already used code"
      });
    }

    markCodeAsUsed(cleanCode, found.file, {
      name,
      mobile,
      purchaseSource
    });

    res.json({
      success: true,
      message: "Product verified successfully"
    });

  } catch (err) {
    console.error("VERIFY ERROR:", err);
    res.status(500).json({
      success: false,
      message: "Server error"
    });
  }
});

/* ---------------- START SERVER ---------------- */
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
