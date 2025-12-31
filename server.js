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

/* ---------------- UTIL: LOAD ALL CODES ---------------- */
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
    const data = XLSX.utils.sheet_to_json(sheet);

    data.forEach(row => {
      if (!row.code) return;

      allCodes.push({
        code: String(row.code).trim(),
        used: row.used === "YES",
        file: file
      });
    });
  });

  return allCodes;
}

/* ---------------- UTIL: MARK CODE AS USED ---------------- */
function markCodeAsUsed(code, fileName) {
  const filePath = path.join(__dirname, "codes", fileName);

  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet);

  data.forEach(row => {
    if (String(row.code).trim() === code) {
      row.used = "YES";
    }
  });

  const updatedSheet = XLSX.utils.json_to_sheet(data);
  workbook.Sheets[sheetName] = updatedSheet;
  XLSX.writeFile(workbook, filePath);
}

/* ---------------- ROUTES ---------------- */

/* HOME (optional) */
app.get("/", (req, res) => {
  res.send("Code Verification Server Running");
});

/* VERIFY CODE */
app.post("/verify", (req, res) => {
  const { code, name, mobile, source } = req.body;

  if (!code) {
    return res.json({
      success: false,
      message: "Code is required"
    });
  }

  const allCodes = loadAllCodes();

  const found = allCodes.find(
    c => c.code === String(code).trim() && c.used === false
  );

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
    purchase_source: found.file.replace(".xlsx", "")
  });
});

/* ---------------- SERVER START ---------------- */
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
