import { getGraphToken } from "../auth/msalClient";

// 取得使用者 (保留給全域搜尋使用)
export async function fetchEntraUsers() {
  console.log("嘗試取得 Graph Token (Users)...");
  const token = await getGraphToken();
  if (!token) throw new Error("未登入或 Token 獲取失敗");

  const url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,department,jobTitle,officeLocation&$top=999";

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  });

  if (!response.ok) throw new Error(`Graph Users 載入失敗: ${response.status}`);
  const data = await response.json();
  return data.value || [];
}

// 取得組織群組
export async function fetchEntraGroups() {
  console.log("嘗試取得 Graph Token (Groups)...");
  const token = await getGraphToken();
  if (!token) throw new Error("未登入");

  const url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName&$top=999";

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  });

  if (!response.ok) throw new Error(`Graph Groups 載入失敗: ${response.status}`);
  const data = await response.json();
  return data.value || [];
}

// 🔥 新增：取得特定群組的成員
export async function fetchGroupMembers(groupId) {
  const token = await getGraphToken();
  if (!token) throw new Error("未登入");

  // 只抓取需要的欄位
  const url = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,jobTitle`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  });

  if (!response.ok) {
    console.warn(`無法讀取群組 ${groupId} 的成員: ${response.statusText}`);
    return [];
  }

  const data = await response.json();
  // 只回傳 User 類型的成員
  return (data.value || []).filter(m => m['@odata.type'] === '#microsoft.graph.user');
}