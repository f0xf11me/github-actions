const XLSX = require("xlsx");
const fs = require("fs");

const issueNumber = process.argv[2];
const issueTitle = process.argv[3];

const dirPath = ".github/excel";
const filePath = `${dirPath}/data.xlsx`;

if (!fs.existsSync(dirPath)) {
  fs.mkdirSync(dirPath, { recursive: true });
}

let workbook;

try {
  workbook = XLSX.readFile(filePath);
} catch (e) {
  workbook = XLSX.utils.book_new();
}

const sheetName = "Issues";
let sheet = workbook.Sheets[sheetName];

if (!sheet) {
  // 初期化（ヘッダー追加）
  sheet = XLSX.utils.aoa_to_sheet([["Issue Number", "Title"]]);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
}

// シートのデータ取得
const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

// 次の空行に追加
sheetData.push([issueNumber, issueTitle]);

// シート再生成
const newSheet = XLSX.utils.aoa_to_sheet(sheetData);
workbook.Sheets[sheetName] = newSheet;

// 書き込み
XLSX.writeFile(workbook, filePath);

console.log(`✅ Issue #${issueNumber} を Excel に追記しました`);
