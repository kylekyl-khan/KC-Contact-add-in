const orgMembers = require("../../data/orgMembers.generated.json");
const { orgTreeConfig } = require("./orgTreeConfig");
const { getGroupMembers } = require("../services/graph/groupMembers");

// 切換 mock 資料或 Graph 真實查詢：改為 true 即會走 Graph 呼叫
const USE_GRAPH = false;

// 簡易快取，避免重複處理同一個 groupId
const groupCache = new Map();
let searchCache = null;

/**
 * 依據 Entra groupId 回傳該群組的成員清單
 * 現階段讀取 data/orgMembers.generated.json；未來接入 Microsoft Graph 時，
 * 可以在此改用 @microsoft/microsoft-graph-client 並搭配 OAuth token 呼叫 /groups/{id}/members。
 * @param {string} groupId
 * @returns {Promise<Array<{id:string,name:string,title:string,email:string}>>}
 */
async function getMembersByGroupId(groupId) {
  if (!groupId) return [];
  if (groupCache.has(groupId)) {
    return groupCache.get(groupId);
  }

  // 目前直接讀取預先產生的 JSON；Outlook taskpane 會被 webpack 打包，
  // 因此不用額外處理檔案 IO。
  const members = orgMembers[groupId] || [];
  groupCache.set(groupId, members);
  return members;
}

/**
 * 單一介面切換使用 Graph 或離線 snapshot
 * @param {string} groupId
 * @returns {Promise<Array<{id,name,title,email}>>}
 */
async function loadMembersForGroup(groupId) {
  if (!groupId) return [];

  if (USE_GRAPH) {
    if (groupCache.has(groupId)) {
      return groupCache.get(groupId);
    }
    const members = await getGroupMembers(groupId);
    groupCache.set(groupId, members);
    return members;
  }

  return getMembersByGroupId(groupId);
}

/**
 * 為搜尋功能整合所有葉節點的成員，並附上組織路徑
 * @returns {Promise<Array<{id,name,title,email,path}>>}
 */
async function getAllMembersWithPath() {
  if (searchCache) return searchCache;

  const result = [];
  const walk = async (nodes, path = []) => {
    for (const node of nodes) {
      const currentPath = [...path, node.name];
      if (node.groupId) {
        const members = await loadMembersForGroup(node.groupId);
        members.forEach(m => {
          result.push({ ...m, path: currentPath.join(" / ") });
        });
      }
      if (Array.isArray(node.employees)) {
        node.employees.forEach(emp => {
          result.push({ ...emp, path: currentPath.join(" / ") });
        });
      }
      if (Array.isArray(node.children)) {
        await walk(node.children, currentPath);
      }
    }
  };

  await walk(orgTreeConfig);
  searchCache = result;
  return result;
}

module.exports = {
  getMembersByGroupId,
  getAllMembersWithPath,
  loadMembersForGroup,
};
