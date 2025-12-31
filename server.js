const express = require("express");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

/* ---------------- MIDDLEWARE ---------------- */
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static("public"));

app.use((req, res, next) => {
  res.setHeader("X-Frame-Options", "ALLOWALL");
  next();
});

/* ---------------- LOAD ALL CODES FROM ALL FILES ---------------- */
function loadAllCodes() {
  const folderPath = path.join(__dirname, "codes");
  if (!fs.existsSync(folderPath)) return [];

  const files = fs.readdirSync(folderPath);
  let allCodes = [];

  files.forEach(file => {
    if (!file.endsWith(".xlsx")) return;

    const filePath = path.join(folderPath, file);
    const workbook = XLSX.readFile(filePath);
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

/* ---------------- MARK CODE AS USED ---------------- */
function markCodeAsUsed(code, fileName) {
  const filePath = path.join(__dirname, "codes", fileName);

  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  let updated = false;

  rows.forEach(row => {
    if (String(row.code).trim() === code) {
      row.used = "YES";
      updated = true;
    }
  });

  if (!updated) return;

  workbook.Sheets[sheetName] = XLSX.utils.json_to_sheet(rows);
  XLSX.writeFile(workbook, filePath);
}

/* ---------------- ROUTES ---------------- */

app.get("/", (req, res) => {
  res.send("Code Verification Server Running");
});

/* DEBUG ROUTE (remove later if needed) */
app.get("/debug", (req, res) => {
  res.json(loadAllCodes().slice(0, 10));
});

/* VERIFY CODE */
app.post("/verify", (req, res) => {
  const inputCode = String(req.body.code || "").trim();

  if (!inputCode) {
    return res.json({
      success: false,
      message: "Code is required"
    });
  }

  const allCodes = loadAllCodes();

  console.log("TOTAL CODES:", allCodes.length);
  console.log("INPUT CODE:", inputCode);

  const found = allCodes.find(
    c => c.code === inputCode && c.used === false
  );

  if (!found) {
    return res.json({
      success: false,
      message: "Invalid or already used code"
    });
  }

  markCodeAsUsed(inputCode, found.file);

  res.json({
    success: true,
    message: "Product verified successfully",
    purchase_source: found.file.replace(".xlsx", "")
  });
});

/* ---------------- START SERVER ---------------- */
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
