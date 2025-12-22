import express from "express";
import fs from "fs";
import XLSX from "xlsx";
import path from "path";
import cors from "cors";


const app = express();
app.use(cors());
app.use(express.json());

const EXCEL_PATH = path.join(process.cwd(), "codes.xlsx");

function readWorkbook() {
  const workbook = XLSX.readFile(EXCEL_PATH);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return { workbook, data: XLSX.utils.sheet_to_json(sheet) };
}

function writeWorkbook(workbook, data) {
  const newSheet = XLSX.utils.json_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = newSheet;
  XLSX.writeFile(workbook, EXCEL_PATH);
}

// Root test
app.get("/", (req, res) => {
  res.send("✅ Code Verification API is running!");
});

// Verify route
app.get("/verify", (req, res) => {
  const { code, name, mobile } = req.query;
  if (!code) return res.json({ success: false, message: "Code missing" });

  const { workbook, data } = readWorkbook();
  const record = data.find((r) => r.code?.trim() === code.trim());

  if (!record) {
    return res.json({ success: false, message: "❌ Invalid code." });
  }

  if (record.status === "used") {
    return res.json({ success: false, message: "⚠️ Code already used." });
  }

  // Mark as used
  record.status = "used";
  record.name = name || "";
  record.mobile = mobile || "";
  record.date_used = new Date().toISOString();

  writeWorkbook(workbook, data);

  return res.json({ success: true, message: "✅ Code verified successfully!" });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

app.get("/download", (req, res) => {
  const filePath = path.resolve("./codes.xlsx"); // adjust name/path if needed
  res.download(filePath, "codes.xlsx", (err) => {
    if (err) {
      console.error("Error downloading file:", err);
      res.status(500).send("Could not download file.");
    }
  });
});

