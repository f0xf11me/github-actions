const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const fetch = require('node-fetch');

// 引数
const eventName = process.argv[2]; // pull_request
const prBody = process.argv[3];    // "#18" など含むPR本文
const isMerged = process.argv[4] === 'true'; // 'true' or 'false'

const token = process.env.GITHUB_TOKEN;
const repo = process.env.GITHUB_REPOSITORY;

console.log(`🟢 イベント名: ${eventName}`);
console.log(`🔄 マージ済み: ${isMerged}`);

if (!prBody || !token || !repo || !eventName) {
  console.error("❌ 引数または環境変数が不足しています");
  process.exit(1);
}

// PR本文から `#18` のような Issue 番号を抽出
const match = prBody.match(/#(\d+)/);
if (!match) {
  console.error("❌ PR body に issue 番号 (#xx) が見つかりませんでした");
  process.exit(1);
}

const issueNumber = match[1];

// GitHub API から issue タイトル取得
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

// Excel に「済」記入
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

  // 📍 D列: PR 作成 or 更新時 ／ F列: マージ時
  const targetCol = isMerged ? 5 : 3;

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
