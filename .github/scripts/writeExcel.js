const XLSX = require("xlsx");
const fs = require("fs");

const filePath = "data.xlsx";
let workbook;

try {
  // 既存ファイルがあれば読み込み
  workbook = XLSX.readFile(filePath);
} catch (e) {
  // なければ新規作成
  workbook = XLSX.utils.book_new();
}

const sheetName = "Sheet1";

// 既存シートがあるかチェック
let sheet = workbook.Sheets[sheetName];

// なければ新規作成し追加
if (!sheet) {
  sheet = XLSX.utils.aoa_to_sheet([["Hello from GitHub Actions!"]]);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
} else {
  // A1 に書き込み（シートがある場合）
  sheet["A1"] = { t: "s", v: "Hello from GitHub Actions!" };
  workbook.Sheets[sheetName] = sheet;
}

// 保存
XLSX.writeFile(workbook, filePath);

console.log("✅ Excel にデータを書き込みました:", filePath);
