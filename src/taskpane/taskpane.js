/* global Office, document */
import { fetchEntraUsers, fetchEntraGroups } from "../services/graph/entraUsers";

// === 全域變數 ===
let allUsers = [];
let orgTree = null;
let orgNodeIndex = {};
const CAMPUS_PREFIX_MAP = {
  "KCHC": "新竹校區",
  "KCQS": "青山校區",
  "NJ": "南京校區",
  "KS": "康軒集團",
  "K1": "康軒集團",
  "KKC": "康橋幼兒園"
};

Office.onReady(() => {
  // ... (UI 初始化代碼保持不變) ...
  initializeOrgUI();
});

async function initializeOrgUI() {
  try {
    console.log("🚀 開始初始化...");
    
    // 1. 抓取資料 (加入個別錯誤處理，避免一個失敗全軍覆沒)
    let users = [], groups = [];
    
    try {
      users = await fetchEntraUsers();
      console.log(`✅ 成功抓取使用者: ${users.length} 筆`);
    } catch (e) {
      console.error("❌ 抓取使用者失敗:", e);
    }

    try {
      groups = await fetchEntraGroups();
      console.log(`✅ 成功抓取群組: ${groups.length} 筆`);
      // 🔍 測試印出第一筆群組，確認格式
      if (groups.length > 0) console.log("🔍 群組資料範例:", groups[0]);
    } catch (e) {
      console.error("❌ 抓取群組失敗 (請檢查 API 權限 Group.Read.All):", e);
    }

    if (groups.length === 0) {
      console.warn("⚠️ 沒有群組資料，將無法建立完整樹狀圖。");
    }

    allUsers = users;
    
    // 2. 建立樹狀骨架
    console.log("🌲 正在建立組織樹...");
    orgTree = buildOrgTreeStructure(groups);
    console.log("🌲 樹狀骨架建立完成:", orgTree);

    // 3. 將人員填入
    console.log("👤 正在填入人員...");
    populateUsersIntoTree(users);

    // 4. 渲染 UI (請確保你有這個函式)
    // renderOrgTree(orgTree); 
    console.log("🎉 初始化完成！");

  } catch (e) {
    console.error("💥 初始化發生致命錯誤：", e);
  }
}

// === 渲染 UI (安全版，避開 innerHTML) ===
function renderOrgTree(rootNode) {
  const treeContainer = document.getElementById("org-tree");
  if (!treeContainer) return;
  
  treeContainer.innerHTML = ""; // 清空容器 (這是唯一允許的操作)
  
  // 遞迴渲染函式
  function createTreeNodeElement(node) {
    // 1. 建立容器
    const nodeEl = document.createElement("div");
    nodeEl.className = "tree-node";
    nodeEl.style.marginLeft = "15px"; // 簡單縮排

    // 2. 建立標題列 (包含展開/收合圖示與名稱)
    const titleRow = document.createElement("div");
    titleRow.className = "node-title";
    titleRow.style.cursor = "pointer";
    titleRow.style.padding = "4px";
    
    // 圖示 (使用文字代替 icon 以避免載入問題，或者用 span class)
    const icon = document.createElement("span");
    const hasChildren = node.children && node.children.length > 0;
    icon.textContent = hasChildren ? "📂 " : "📁 ";
    
    // 名稱
    const nameSpan = document.createElement("span");
    nameSpan.textContent = `${node.name} (${node.users.length})`;
    nameSpan.style.fontWeight = node.users.length > 0 ? "bold" : "normal";

    titleRow.appendChild(icon);
    titleRow.appendChild(nameSpan);

    // 3. 點擊事件：展開/收合 或 顯示成員
    titleRow.onclick = (e) => {
      e.stopPropagation();
      // 切換子節點顯示
      if (childrenContainer) {
        const isHidden = childrenContainer.style.display === "none";
        childrenContainer.style.display = isHidden ? "block" : "none";
        icon.textContent = isHidden ? "📂 " : "📁 ";
      }
      // 觸發顯示成員 (呼叫外部函式)
      showContacts(node); 
    };

    nodeEl.appendChild(titleRow);

    // 4. 建立子節點容器
    let childrenContainer = null;
    if (hasChildren) {
      childrenContainer = document.createElement("div");
      childrenContainer.className = "node-children";
      childrenContainer.style.display = "none"; // 預設收合，避免畫面太長
      
      // 遞迴建立子節點
      node.children.forEach(child => {
        childrenContainer.appendChild(createTreeNodeElement(child));
      });
      nodeEl.appendChild(childrenContainer);
    }

    return nodeEl;
  }

  // 開始渲染
  if (rootNode) {
    // 因為 root 包含多個校區，我們直接遍歷 root.children
    rootNode.children.forEach(campus => {
       treeContainer.appendChild(createTreeNodeElement(campus));
    });
  }
}

// 輔助函式：顯示成員 (這部分不需要動 innerHTML，也建議用 DOM API)
function showContacts(node) {
  const listContainer = document.getElementById("contacts-list");
  listContainer.innerHTML = ""; // 清空

  if (!node.users || node.users.length === 0) {
    const emptyMsg = document.createElement("div");
    emptyMsg.textContent = "此群組無成員";
    emptyMsg.style.color = "#888";
    emptyMsg.style.padding = "10px";
    listContainer.appendChild(emptyMsg);
    return;
  }

  node.users.forEach(user => {
    const item = document.createElement("div");
    item.className = "contact-item";
    item.style.padding = "8px";
    item.style.borderBottom = "1px solid #eee";
    item.style.cursor = "pointer";

    // 名稱
    const name = document.createElement("div");
    name.textContent = user.displayName;
    name.style.fontWeight = "bold";

    // Email
    const email = document.createElement("div");
    email.textContent = user.mail || user.userPrincipalName;
    email.style.fontSize = "0.85em";
    email.style.color = "#666";

    item.appendChild(name);
    item.appendChild(email);
    
    // 點擊事件 (加入收件人)
    item.onclick = () => {
        // 這裡呼叫你原本的 addToRecipients 邏輯
        console.log("選取使用者:", user.displayName);
        // addRecipientToSelection(user); // 假設你有這個函式
    };

    listContainer.appendChild(item);
  });
}

function buildOrgTreeStructure(groups) {
  orgNodeIndex = {}; 
  const root = { id: "root", name: "康橋通訊錄", children: [], users: [] };
  
  // 建立校區節點
  const campusNodes = {};
  for (const [prefix, name] of Object.entries(CAMPUS_PREFIX_MAP)) {
    if (!campusNodes[name]) {
      const node = { id: `campus-${prefix}`, name: name, children: [], users: [], type: 'campus' };
      campusNodes[name] = node;
      root.children.push(node);
    }
  }

  // 解析群組 (加強 Debug)
  let parsedCount = 0;
  let parsedGroups = groups.map(g => {
    // 嘗試解析 "K10010.康軒經管會議" 或 "K10010 康軒經管會議"
    // Regex 解釋：
    // ^([A-Z0-9]+) -> 開頭是英數字 (Code)
    // [\.\-_\s]+   -> 中間是 點、減號、底線或空白
    // (.+)$        -> 後面是 名稱
    const match = g.displayName && g.displayName.match(/^([A-Z0-9]+)[\.\-_\s]+(.+)$/);
    
    if (match) {
      parsedCount++;
      return { original: g, code: match[1], name: match[2].trim() };
    } else {
      // 若解析失敗，可在這裡 log 看看為什麼失敗
      // console.log("無法解析群組名稱:", g.displayName); 
      return null; 
    }
  }).filter(g => g !== null);

  console.log(`📊 解析成功群組數: ${parsedCount} / ${groups.length}`);

  parsedGroups.sort((a, b) => a.code.length - b.code.length || a.code.localeCompare(b.code));

  // 建立節點索引
  parsedGroups.forEach(pg => {
    orgNodeIndex[pg.code] = { id: pg.code, name: pg.name, children: [], users: [] };
  });

  // 建立層級
  parsedGroups.forEach(pg => {
    const currentNode = orgNodeIndex[pg.code];
    let parentFound = false;

    // 往回找父節點 (e.g. KCHC100101 -> KCHC1001 -> KCHC10)
    for (let i = pg.code.length - 1; i >= 2; i--) {
      const parentCode = pg.code.substring(0, i);
      if (orgNodeIndex[parentCode]) {
        orgNodeIndex[parentCode].children.push(currentNode);
        parentFound = true;
        break;
      }
    }

    if (!parentFound) {
      // 找不到父群組，嘗試歸類到校區
      let assigned = false;
      for (const [prefix, campusName] of Object.entries(CAMPUS_PREFIX_MAP)) {
        if (pg.code.startsWith(prefix)) {
          campusNodes[campusName].children.push(currentNode);
          assigned = true;
          break;
        }
      }
      // 如果連校區都沒有，這是一個孤兒節點 (Orphan)，暫時掛在根目錄底下以便除錯
      if (!assigned) {
         // root.children.push(currentNode); // 解開註解可顯示未分類群組
      }
    }
  });

  return root;
}

function populateUsersIntoTree(users) {
  let mappedCount = 0;
  users.forEach(u => {
    if (!u.department) return;
    
    // 嘗試從 department 字串 (e.g. "KCHC1010.新竹教務處") 抓出代碼
    const match = u.department.match(/^([A-Z0-9]+)/);
    if (match) {
      const code = match[1];
      if (orgNodeIndex[code]) {
        orgNodeIndex[code].users.push(u);
        mappedCount++;
      }
    }
  });
  console.log(`📌 成功定位人員: ${mappedCount} / ${users.length}`);
}