const express = require("express");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

/* ---------------- HARD CORS FIX (SHOPIFY SAFE) ---------------- */
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
    return res.sendStatus(200); // PRE-FLIGHT RESPONSE
  }

  next();
});

/* ---------------- MIDDLEWARE ---------------- */
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

/* ---------------- BASIC TEST ---------------- */
app.get("/", (req, res) => {
  res.send("SERVER OK");
});

/* ---------------- LOAD ALL CODES ---------------- */
function loadAllCodes() {
  const folderPath = path.join(__dirname, "codes");
  if (!fs.existsSync(folderPath)) return [];

  const files = fs.readdirSync(folderPath);
  let allCodes = [];

  files.forEach(file => {
    if (!file.endsWith(".xlsx")) return;

    const workbook = XLSX.readFile(path.join(folderPath, file));
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
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

/* ---------------- MARK CODE USED ---------------- */
function markCodeAsUsed(code, file) {
  const filePath = path.join(__dirname, "codes", file);
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(
    workbook.Sheets[sheetName],
    { defval: "" }
  );

  rows.forEach(row => {
    if (String(row.code).trim() === code) {
      row.used = "YES";
    }
  });

  workbook.Sheets[sheetName] = XLSX.utils.json_to_sheet(rows);
  XLSX.writeFile(workbook, filePath);
}

/* ---------------- VERIFY ROUTE ---------------- */
app.post("/verify", (req, res) => {
  try {
    const code = String(req.body.code || "").trim();
    if (!code) {
      return res.json({ success: false, message: "Code required" });
    }

    const allCodes = loadAllCodes();
    const found = allCodes.find(c => c.code === code && !c.used);

    if (!found) {
      return res.json({
        success: false,
        message: "Invalid or already used code"
      });
    }

    markCodeAsUsed(code, found.file);

    res.json({
      success: true,
      message: "Product verified successfully",
      source: found.file.replace(".xlsx", "")
    });
  } catch (err) {
    console.error("VERIFY ERROR:", err);
    res.status(500).json({ success: false, message: "Server error" });
  }
});

/* ---------------- START ---------------- */
app.listen(PORT, () => {
  console.log("SERVER STARTED ON PORT", PORT);
});
