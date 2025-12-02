import { getGraphToken } from "../auth/msalClient";

// 取得使用者 (維持不變，確保抓取 department)
export async function fetchEntraUsers() {
  console.log("嘗試取得 Graph Token (Users)...");
  const token = await getGraphToken();
  if (!token) throw new Error("未登入或 Token 獲取失敗");

  // 確保包含 department
  const url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,department,jobTitle,officeLocation&$top=999";

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  });

  if (!response.ok) throw new Error(`Graph Users 載入失敗: ${response.status}`);
  const data = await response.json();
  return data.value || [];
}

// 🔥 新增：取得組織群組 (用來建立樹狀骨架)
export async function fetchEntraGroups() {
  console.log("嘗試取得 Graph Token (Groups)...");
  const token = await getGraphToken();
  if (!token) throw new Error("未登入");

  // 抓取群組，只需 id 和 displayName
  const url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName&$top=999";

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
  });

  if (!response.ok) throw new Error(`Graph Groups 載入失敗: ${response.status}`);
  const data = await response.json();
  return data.value || [];
}