#!/usr/bin/env node

const fs = require("fs");
const path = require("path");
const readline = require("readline");
const unzipper = require("unzipper");
const XLSX = require("xlsx");

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

const cwd = process.cwd();

// 1️⃣ 找 zip 文件
const zips = fs.readdirSync(cwd).filter(f => f.endsWith(".zip"));

if (zips.length === 0) {
  console.log("当前目录没有 zip 文件");
  process.exit();
}

console.log("发现以下压缩包：");
zips.forEach((z, i) => console.log(`${i + 1}. ${z}`));

rl.question("请选择编号：", async (choice) => {
  const zipFile = zips[choice - 1];
  if (!zipFile) {
    console.log("无效选择");
    process.exit();
  }

  const folderName = path.basename(zipFile, ".zip");
  const targetDir = path.join(cwd, folderName);

  if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir);
  }

  console.log("解压中...");

  await fs.createReadStream(zipFile)
    .pipe(unzipper.Extract({ path: targetDir }))
    .promise();

  console.log("解压完成");

  // 2️⃣ 找 xlsx 文件
  const excelFile = fs.readdirSync(targetDir)
    .find(f => f.endsWith(".xlsx"));

  if (!excelFile) {
    console.log("未找到 xlsx 文件");
    process.exit();
  }

  const excelPath = path.join(targetDir, excelFile);

  // 3️⃣ 构建映射
  const workbook = XLSX.readFile(excelPath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);

  const map = {};

  rows.forEach(row => {
    const date = new Date(row["开票日期"]);
    const amount = Number(row["金额"]).toFixed(2);
    const type = row["发票类型"];

    if (!date || !amount || !type) return;

    const yy = String(date.getFullYear()).slice(2);
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");

    const shortDate = `${yy}${mm}${dd}`;
    map[`${shortDate}_${amount}`] = type;
  });

  // 4️⃣ 重命名 pdf
  const files = fs.readdirSync(targetDir);

  files.forEach(file => {
    if (!file.endsWith(".pdf")) return;

    const match = file.match(/^(\d{6})_(\d+\.\d{2})_(.+)\.pdf$/);
    if (!match) return;

    const [_, date, amount] = match;

    const type = map[`${date}_${amount}`];
    if (!type) {
      console.log("未匹配到类型:", file);
      return;
    }

    const newName = `${date}-${type}-${amount}.pdf`;

    fs.renameSync(
      path.join(targetDir, file),
      path.join(targetDir, newName)
    );

    console.log(`${file} → ${newName}`);
  });

  console.log("处理完成！");
  rl.close();
});