// 必要なパッケージ: npm install xlsx node-fetch@2
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const fetch = require('node-fetch');

const prBody = process.argv[2];
const token = process.env.GITHUB_TOKEN;
const repo = process.env.GITHUB_REPOSITORY;

if (!prBody || !token || !repo) {
  console.error("❌ 引数または環境変数が不足しています");
  process.exit(1);
}

// PR 本文から #番号 を取得
const match = prBody.match(/#(\d+)/);
if (!match) {
  console.error("❌ PR body に issue 番号 (#xx) が見つかりませんでした");
  process.exit(1);
}

const issueNumber = match[1];

// GitHub API から Issue タイトルを取得
async function getIssueTitle() {
  const url = `https://api.github.com/repos/${repo}/issues/${issueNumber}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/vnd.github+json',
    },
  });

  if (!res.ok) {
    throw new Error(`GitHub API error: ${res.status}`);
  }

  const json = await res.json();
  return json.title;
}

async function writeToExcel() {
  const issueTitle = await getIssueTitle();

  console.log(`📄 Issue #${issueNumber}: ${issueTitle}`);

  const filePath = path.resolve(__dirname, '../excel/data.xlsx');
  let workbook;
  let worksheet;

  if (fs.existsSync(filePath)) {
    workbook = xlsx.readFile(filePath);
    worksheet = workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = xlsx.utils.book_new();
    worksheet = xlsx.utils.aoa_to_sheet([["Issue Number", "Title"]]);
    xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  }

  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  data.push([`#${issueNumber}`, issueTitle]);

  const newSheet = xlsx.utils.aoa_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = newSheet;
  xlsx.writeFile(workbook, filePath);

  console.log("✅ Excel ファイルに追記しました");
}

writeToExcel().catch(err => {
  console.error("❌ エラー:", err.message);
  process.exit(1);
});
