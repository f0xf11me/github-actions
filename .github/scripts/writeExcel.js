const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const fetch = require('node-fetch');

// 引数：event名, PR本文
const eventName = process.argv[2];
const prBody = process.argv[3];
const token = process.env.GITHUB_TOKEN;
const repo = process.env.GITHUB_REPOSITORY;

if (!prBody || !token || !repo || !eventName) {
  console.error("❌ 引数または環境変数が不足しています");
  process.exit(1);
}

console.log(`📦 GitHub Event: ${eventName}`);

// #番号 を PR body から抽出
const match = prBody.match(/#(\d+)/);
if (!match) {
  console.error("❌ PR body に issue 番号 (#xx) が見つかりませんでした");
  process.exit(1);
}

const issueNumber = match[1];

// GitHub API で issue タイトルを取得
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

// Excel 更新処理
async function markAsDone(issueTitle) {
  const filePath = path.resolve(__dirname, '../excel/data.xlsx');
  const workbook = xlsx.readFile(filePath);
  const sheetName = "issue一覧";
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    throw new Error(`❌ シート '${sheetName}' が見つかりません`);
  }

  const range = xlsx.utils.decode_range(worksheet["!ref"]);
  let updated = false;

  // イベントに応じた列番号を決定（D列 = 3, F列 = 5）
  const targetCol = eventName === 'pull_request_review' ? 5 : 3;

  for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const titleCell = worksheet[xlsx.utils.encode_cell({ r: row, c: 1 })]; // B列

    if (titleCell && titleCell.v === issueTitle) {
      const statusCellAddr = xlsx.utils.encode_cell({ r: row, c: targetCol });
      worksheet[statusCellAddr] = { t: 's', v: '済' };
      updated = true;

      console.log(`✅ '${issueTitle}' に一致：${statusCellAddr} に「済」を書き込みました`);
      break;
    }
  }

  if (!updated) {
    console.log(`⚠️ 一致するタイトル '${issueTitle}' が見つかりませんでした`);
  }

  xlsx.writeFile(workbook, filePath);
}

(async () => {
  try {
    const title = await getIssueTitle();
    console.log(`📝 Issue タイトル: ${title}`);
    await markAsDone(title);
  } catch (err) {
    console.error("❌ エラー:", err.message);
    process.exit(1);
  }
})();
