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
let sheet = workbook.Sheets[sheetName];

// シートが存在しない場合、新しく作成して追加する
if (!sheet) {
  sheet = XLSX.utils.aoa_to_sheet([["Hello from GitHub Actions!"]]);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
} else {
  sheet["A1"] = { t: "s", v: "Hello from GitHub Actions!" };
  workbook.Sheets[sheetName] = sheet;
}

// 最後に書き込む
XLSX.writeFile(workbook, filePath);

console.log("✅ Excel にデータを書き込みました:", filePath);
