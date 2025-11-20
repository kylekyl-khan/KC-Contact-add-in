/*
 * 從 exportGroups_2025-11-19.csv 抽取 orgTreeConfig 中的 groupId，
 * 並產生 data/orgMembers.generated.json（目前為示範假資料）。
 * 未來若要接 Microsoft Graph，可以改為呼叫 Graph API 取得真實成員，再輸出相同格式。
 */
const fs = require("fs");
const path = require("path");
const { orgTreeConfig } = require("../src/data/orgTreeConfig");

const CSV_PATH = path.resolve(__dirname, "../data/exportGroups_2025-11-19.csv");
const OUTPUT_PATH = path.resolve(__dirname, "../data/orgMembers.generated.json");

function readCsv() {
  const raw = fs.readFileSync(CSV_PATH, "utf-8").replace(/^\uFEFF/, "");
  const lines = raw.split(/\r?\n/).filter(Boolean);
  const header = lines.shift();
  const headers = header.split(",");

  return lines.map(line => {
    const cols = line.split(",");
    const record = {};
    headers.forEach((key, idx) => {
      record[key] = cols[idx];
    });
    return record;
  });
}

function collectGroupIds(nodes, list = []) {
  nodes.forEach(node => {
    if (node.groupId) list.push(node.groupId);
    if (Array.isArray(node.children)) collectGroupIds(node.children, list);
  });
  return list;
}

function deriveMembersFromGroup(group) {
  const displayName = group.displayName || group.id;
  const [, orgName = displayName] = displayName.split(".");
  const mailPrefix = (group.mail || `${group.id}@example.com`).split("@")[0];

  return [1, 2, 3].map(idx => ({
    id: `${group.id}-member-${idx}`,
    name: `${orgName} 成員${idx}`,
    title: "示範資料",
    email: `${mailPrefix}.member${idx}@kcis.generated`
  }));
}

function main() {
  const csvRecords = readCsv();
  const targetIds = collectGroupIds(orgTreeConfig);
  const generated = {};

  targetIds.forEach(groupId => {
    const group = csvRecords.find(rec => rec.id === groupId);
    if (!group) {
      console.warn(`找不到 groupId: ${groupId}，請確認 CSV 是否同步。`);
      generated[groupId] = [];
      return;
    }
    generated[groupId] = deriveMembersFromGroup(group);
  });

  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(generated, null, 2), "utf-8");
  console.log(`已產生 ${OUTPUT_PATH}`);
}

main();
