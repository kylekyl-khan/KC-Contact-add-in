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
    throw error; // 將錯誤往上拋，讓呼叫者知道失敗了
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

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Graph Users 載入失敗: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    console.log(`✅ 成功載入使用者: ${data.value ? data.value.length : 0} 筆`);
    return data.value || [];

  } catch (e) {
    console.error("❌ fetchEntraUsers 流程失敗:", e);
    // 這裡我們不 throw，避免影響主程式其他部分 (如群組顯示)
    // 回傳空陣列，讓 UI 繼續運作
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

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph Groups 載入失敗: ${response.status} - ${errorText}`);
    }
    
    const data = await response.json();
    return data.value || [];

  } catch (e) {
    console.error("❌ fetchEntraGroups 流程失敗:", e);
    throw e; // 群組是核心功能，失敗需要拋出錯誤給 taskpane 處理
  }
}

// 🔥 取得特定群組的成員 (這是你點擊組織時會呼叫的)
export async function fetchGroupMembers(groupId) {
  console.log(`🔍 正在載入群組成員 (GroupID: ${groupId})...`);
  
  try {
    const token = await getTokenWithLog("fetchGroupMembers");

    // 只抓取需要的欄位
    const url = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,jobTitle`;

    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    });

    if (!response.ok) {
      console.warn(`⚠️ 無法讀取群組 ${groupId} 的成員: ${response.status} ${response.statusText}`);
      return [];
    }

    const data = await response.json();
    const members = (data.value || []).filter(m => m['@odata.type'] === '#microsoft.graph.user');
    
    console.log(`✅ 群組成員載入完成: ${members.length} 位`);
    return members;

  } catch (e) {
    console.error(`❌ fetchGroupMembers 失敗 (GroupId: ${groupId}):`, e);
    // 回傳空陣列，避免 UI 崩潰，但會在 Console 留下記錄
    return [];
  }
}