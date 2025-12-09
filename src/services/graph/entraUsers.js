import { getGraphToken } from "../auth/msalClient";

/**
 * 通用的錯誤處理與 Log 輔助函式
 */
async function getTokenWithLog(sourceName) {
  try {
    console.log(`🔑 [${sourceName}] 正在呼叫 getGraphToken()...`);
    const token = await getGraphToken();
    
    if (!token) {
      console.error(`❌ [${sourceName}] Token 為 null 或 undefined`);
      throw new Error("無法取得存取權杖 (Token is null)");
    }

    console.log(`✅ [${sourceName}] 成功取得 Token`);
    return token;
  } catch (error) {
    console.error(`💥 [${sourceName}] 取得 Token 時發生錯誤:`, error);
    throw error;
  }
}

// 取得使用者 (保留給全域搜尋使用)
export async function fetchEntraUsers() {
  console.log("🚀 entraUsers.js: 開始執行 fetchEntraUsers...");
  try {
    const token = await getTokenWithLog("fetchEntraUsers");
    const url = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,department,jobTitle,officeLocation&$top=999";

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    });

    if (!response.ok) throw new Error(`Graph Users 載入失敗: ${response.status}`);
    const data = await response.json();
    console.log(`✅ 成功載入使用者: ${data.value ? data.value.length : 0} 筆`);
    return data.value || [];
  } catch (e) {
    console.error("❌ fetchEntraUsers 流程失敗:", e);
    return [];
  }
}

// 取得組織群組
export async function fetchEntraGroups() {
  console.log("🚀 entraUsers.js: 開始執行 fetchEntraGroups...");
  try {
    const token = await getTokenWithLog("fetchEntraGroups");
    const url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mail&$top=999";

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    });

    if (!response.ok) throw new Error(`Graph Groups 載入失敗: ${response.status}`);
    const data = await response.json();
    return data.value || [];
  } catch (e) {
    console.error("❌ fetchEntraGroups 流程失敗:", e);
    throw e;
  }
}

// 🔥 修正版：取得特定群組的成員 (使用 transitiveMembers 以支援巢狀群組)
export async function fetchGroupMembers(groupId) {
  console.log(`🔍 正在載入群組成員 (GroupID: ${groupId})...`);
  
  try {
    const token = await getTokenWithLog("fetchGroupMembers");

    // 🔥 關鍵修改：使用 transitiveMembers 來展開巢狀群組，確保能抓到子群組內的人
    // 同時過濾掉不需要的物件類型，只保留 user
    const url = `https://graph.microsoft.com/v1.0/groups/${groupId}/transitiveMembers?$select=id,displayName,mail,userPrincipalName,jobTitle,department`;

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    });

    if (!response.ok) {
      console.warn(`⚠️ 無法讀取群組 ${groupId} 的成員: ${response.status} ${response.statusText}`);
      return [];
    }

    const data = await response.json();
    
    // Log 原始長度以便除錯
    if (data.value) {
        console.log(`📦 API 回傳原始筆數: ${data.value.length}`);
    }

    // 只回傳 User 類型的成員 (排除 Group, Device 等其他物件)
    const members = (data.value || []).filter(m => m['@odata.type'] === '#microsoft.graph.user');
    
    console.log(`✅ 成功解析使用者成員: ${members.length} 位`);
    return members;

  } catch (e) {
    console.error(`❌ fetchGroupMembers 失敗 (GroupId: ${groupId}):`, e);
    return [];
  }
}