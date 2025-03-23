const XLSX = require("xlsx");
const fs = require("fs");

// 保存先ディレクトリとファイルパス
const dirPath = ".github/excel";
const filePath = `${dirPath}/data.xlsx`;

// ディレクトリがなければ作成
if (!fs.existsSync(dirPath)) {
  fs.mkdirSync(dirPath, { recursive: true });
}

let workbook;

try {
  workbook = XLSX.readFile(filePath);
} catch (e) {
  workbook = XLSX.utils.book_new();
}

const sheetName = "Sheet1";

// 既存シートがあるか確認
let sheet = workbook.Sheets[sheetName];

if (!sheet) {
  sheet = XLSX.utils.aoa_to_sheet([["Hello from GitHub Actions!"]]);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
} else {
  sheet["A1"] = { t: "s", v: "Hello from GitHub Actions!" };
  workbook.Sheets[sheetName] = sheet;
}

XLSX.writeFile(workbook, filePath);
console.log("✅ Excel に書き込みました:", filePath);
