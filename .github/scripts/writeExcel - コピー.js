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
  // シートがない場合は新規作成して A1 に書く
  sheet = XLSX.utils.aoa_to_sheet([["Hello from GitHub Actions!"]]);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
} else {
  // シートがある場合は A1 が空かどうかを確認
  const cellA1 = sheet["A1"];
  if (!cellA1 || !cellA1.v) {
    // A1 が空ならそこに書く
    sheet["A1"] = { t: "s", v: "Hello from GitHub Actions!" };
  } else {
    // A1 にすでに値があれば、A2 に書く
    sheet["A2"] = { t: "s", v: "Hello from GitHub Actions!" };
  }
  workbook.Sheets[sheetName] = sheet;
}

XLSX.writeFile(workbook, filePath);
console.log("✅ Excel に書き込みました:", filePath);
