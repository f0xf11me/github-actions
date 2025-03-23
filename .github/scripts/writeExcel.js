const XLSX = require("xlsx");
const fs = require("fs");

const filePath = "data.xlsx";
let workbook;
try {
  workbook = XLSX.readFile(filePath);
} catch (e) {
  workbook = XLSX.utils.book_new();
}

const sheetName = "Sheet1";
let sheet = workbook.Sheets[sheetName] || XLSX.utils.aoa_to_sheet([["A1"]]);

sheet["A1"] = { t: "s", v: "Hello from GitHub Actions!" };

workbook.Sheets[sheetName] = sheet;
XLSX.writeFile(workbook, filePath);

console.log("Excel にデータを書き込みました:", filePath);
