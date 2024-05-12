const express = require("express");
const XLSX = require("xlsx");
const app = express();
const port = 3000;

app.get("/xlsx-to-json", (req, res) => {
  const jsonData = readXLSXFile("./files/Duplex_A_20110907.xlsx");
  res.json(jsonData);
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});

function readXLSXFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // Read the first sheet
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);
  return jsonData;
}
