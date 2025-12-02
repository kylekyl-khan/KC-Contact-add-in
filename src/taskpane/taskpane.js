/* global Office, document */
import { fetchEntraUsers, fetchEntraGroups, fetchGroupMembers } from "../services/graph/entraUsers";

// === 全域變數 ===
let allUsers = []; // 暫存全域使用者供搜尋用
let orgTree = null;
let orgNodeIndex = {};
let selectedRecipients = [];

// 定義校區前綴對照表
const CAMPUS_PREFIX_MAP = {
  "KCXG": "秀岡校區",
  "KCQS": "青山校區",
  "KCHC": "新竹校區",
  // "NJ": "南京校區", // 已移除
  "KS": "康軒集團",
  "K1": "康軒集團",
  "KKC": "康橋幼兒園"
};

Office.onReady(() => {
  const sideloadMsg = document.getElementById("sideload-msg");
  const appBody = document.getElementById("app-body");
  if (sideloadMsg) sideloadMsg.style.display = "none";
  if (appBody) {
    appBody.style.display = "flex";
    appBody.style.flexDirection = "column";
  }
  initializeOrgUI();
});

async function initializeOrgUI() {
  try {
    console.log("🚀 開始初始化 (API 模式)...");
    
    // 1. 抓取群組 (這是最優先的)
    let groups = [];
    try {
      groups = await fetchEntraGroups();
      console.log(`✅ 成功抓取群組: ${groups.length} 筆`);
    } catch (e) {
      console.error("❌ 抓取群組失敗:", e);
      throw e; // 群組失敗就無法繼續
    }

    // 2. 建立樹狀骨架
    console.log("🌲 建立組織樹...");
    orgTree = buildOrgTreeStructure(groups);

    // 3. 渲染 UI (使用者此時已經可以看到組織樹)
    console.log("🎨 渲染介面...");
    renderOrgTree(orgTree); 
    
    // 4. 【優化修改】將背景抓取改為 await 串行執行
    // 這樣可以確保它絕對不會跟上面的 fetchEntraGroups 或使用者的點擊操作撞車
    // 雖然叫做"背景"，但為了穩定性，我們讓它乖乖排隊
    try {
        console.log("⏳ 開始載入全域使用者清單 (搜尋用)...");
        const users = await fetchEntraUsers();
        allUsers = users;
        console.log(`✅ 全域使用者清單載入完成: ${users.length} 筆`);
    } catch (e) {
        console.warn("⚠️ 無法載入全域使用者 (不影響樹狀圖功能):", e);
    }

    console.log("🎉 初始化全部完成！系統就緒。");
    setupEventHandlers();

  } catch (e) {
    console.error("💥 初始化錯誤：", e);
    const appBody = document.getElementById("app-body");
    if (appBody) appBody.innerHTML = `<div style="color:red; padding:20px;">初始化錯誤: ${e.message}</div>`;
  }
}

// === 核心邏輯：建立樹狀骨架 ===
function buildOrgTreeStructure(groups) {
  orgNodeIndex = {}; 
  const root = { id: "root", name: "康橋通訊錄", children: [], users: [] };
  
  const campusNodes = {};
  for (const [prefix, name] of Object.entries(CAMPUS_PREFIX_MAP)) {
    if (!campusNodes[name]) {
      const node = { 
          id: `campus-${prefix}`, 
          name: name, 
          children: [], 
          users: [], 
          type: 'campus',
          membersLoaded: true // 校區節點視為已載入(因為它只是容器)
      };
      campusNodes[name] = node;
      root.children.push(node);
    }
  }

  // 解析群組
  let parsedGroups = groups.map(g => {
    const match = g.displayName && g.displayName.match(/^([A-Z0-9]+)[\.\-_\s]+(.+)$/);
    if (match) {
      return { 
        original: g, 
        code: match[1], 
        name: match[2].trim() 
      };
    }
    return null; 
  }).filter(g => g !== null);

  parsedGroups.sort((a, b) => a.code.length - b.code.length || a.code.localeCompare(b.code));

  parsedGroups.forEach(pg => {
    orgNodeIndex[pg.code] = { 
        id: pg.code, // 這是樹的 ID
        name: pg.name, 
        children: [], 
        users: [], 
        original: pg.original, // 保留原始 Graph 資料以便後續查詢 ID
        membersLoaded: false,  // 標記：成員尚未載入
        isLoading: false       // 🔥 新增：標記是否正在載入中 (防止連點)
    };
  });

  parsedGroups.forEach(pg => {
    const currentNode = orgNodeIndex[pg.code];
    let parentFound = false;

    for (let i = pg.code.length - 1; i >= 2; i--) {
      const parentCode = pg.code.substring(0, i);
      if (orgNodeIndex[parentCode]) {
        orgNodeIndex[parentCode].children.push(currentNode);
        parentFound = true;
        break;
      }
    }

    if (!parentFound) {
      for (const [prefix, campusName] of Object.entries(CAMPUS_PREFIX_MAP)) {
        if (pg.code.startsWith(prefix)) {
          campusNodes[campusName].children.push(currentNode);
          break;
        }
      }
    }
  });

  return root;
}

// === 渲染 UI (支援 Lazy Loading) ===
function renderOrgTree(rootNode) {
  const treeContainer = document.getElementById("org-tree");
  if (!treeContainer) return;
  treeContainer.innerHTML = ""; 
  
  function createTreeNodeElement(node) {
    const nodeEl = document.createElement("div");
    nodeEl.className = "tree-node";
    nodeEl.style.marginLeft = "15px";

    const titleRow = document.createElement("div");
    titleRow.className = "node-title";
    titleRow.style.cursor = "pointer";
    titleRow.style.padding = "4px";
    titleRow.style.display = "flex";
    titleRow.style.alignItems = "center";
    
    // Icon
    const icon = document.createElement("span");
    const hasChildren = node.children && node.children.length > 0;
    icon.textContent = hasChildren ? "📁 " : "🔹 ";
    icon.style.marginRight = "5px";
    
    // Name
    const nameSpan = document.createElement("span");
    nameSpan.textContent = node.name; 
    
    // 如果是群組節點且未載入，顯示灰色
    if (!node.membersLoaded && node.original) {
        nameSpan.style.color = "#555";
    }

    titleRow.appendChild(icon);
    titleRow.appendChild(nameSpan);

    // 🔥 點擊事件：Lazy Load 成員 (包含防連點機制)
    titleRow.onclick = async (e) => {
      e.stopPropagation();

      // 1. 如果正在載入中，直接忽略點擊 (防止 interaction_in_progress)
      if (node.isLoading) {
          console.log("⏳ 正在載入中，請稍候...");
          return;
      }

      // 2. 展開/收合子節點 (視覺效果)
      if (childrenContainer) {
        const isHidden = childrenContainer.style.display === "none";
        childrenContainer.style.display = isHidden ? "block" : "none";
        icon.textContent = isHidden ? "📂 " : "📁 ";
      }

      // 3. 如果是群組節點，且還沒載入成員 -> 去 API 抓！
      if (node.original && !node.membersLoaded) {
          // 鎖定狀態
          node.isLoading = true;
          
          nameSpan.textContent = `${node.name} (載入中...)`;
          nameSpan.style.color = "blue";
          
          try {
              // 這裡會觸發 Graph API 呼叫
              const members = await fetchGroupMembers(node.original.id);
              node.users = members;
              node.membersLoaded = true;
              
              // 更新顯示
              nameSpan.textContent = `${node.name} (${members.length})`;
              nameSpan.style.color = members.length > 0 ? "black" : "#888";
              nameSpan.style.fontWeight = members.length > 0 ? "bold" : "normal";
          } catch (err) {
              console.error("載入成員失敗:", err);
              nameSpan.textContent = `${node.name} (載入失敗)`;
              nameSpan.style.color = "red";
          } finally {
              // 無論成功失敗，都解除鎖定
              node.isLoading = false;
          }
      }

      // 4. 顯示成員列表
      showContacts(node); 
    };

    nodeEl.appendChild(titleRow);

    let childrenContainer = null;
    if (hasChildren) {
      childrenContainer = document.createElement("div");
      childrenContainer.className = "node-children";
      childrenContainer.style.display = "none"; 
      
      node.children.forEach(child => {
        childrenContainer.appendChild(createTreeNodeElement(child));
      });
      nodeEl.appendChild(childrenContainer);
    }

    return nodeEl;
  }

  if (rootNode && rootNode.children) {
    rootNode.children.forEach(campus => {
       treeContainer.appendChild(createTreeNodeElement(campus));
    });
  }
}

// === 顯示成員列表 ===
function showContacts(node) {
  const listContainer = document.getElementById("contacts-list");
  if (!listContainer) return;
  listContainer.innerHTML = ""; 

  const breadcrumb = document.getElementById("breadcrumb");
  if (breadcrumb) breadcrumb.textContent = node.name;
  
  const countSpan = document.getElementById("contacts-count");
  if (countSpan) {
      if (node.membersLoaded) {
        countSpan.textContent = `共 ${node.users.length} 筆`;
      } else {
        countSpan.textContent = "點擊載入...";
      }
  }

  if (!node.users || node.users.length === 0) {
    const emptyMsg = document.createElement("div");
    emptyMsg.textContent = node.membersLoaded ? "此群組無成員" : "請點擊群組標題以載入成員";
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
    item.style.display = "flex";
    item.style.justifyContent = "space-between";
    item.style.alignItems = "center";

    const infoDiv = document.createElement("div");
    const nameDiv = document.createElement("div");
    nameDiv.textContent = user.displayName;
    nameDiv.style.fontWeight = "bold";
    const emailDiv = document.createElement("div");
    emailDiv.textContent = user.mail || user.userPrincipalName;
    emailDiv.style.fontSize = "0.85em";
    emailDiv.style.color = "#666";

    infoDiv.appendChild(nameDiv);
    infoDiv.appendChild(emailDiv);

    const addBtn = document.createElement("button");
    addBtn.textContent = "+";
    addBtn.style.padding = "2px 8px";
    
    item.appendChild(infoDiv);
    item.appendChild(addBtn);
    
    item.onclick = () => addToSelection(user);
    listContainer.appendChild(item);
  });
}

// === 選取清單與其他功能 (維持不變) ===
function addToSelection(user) {
    if (selectedRecipients.find(u => u.id === user.id)) return;
    selectedRecipients.push(user);
    renderSelectionList();
}

function renderSelectionList() {
    const container = document.getElementById("selection-list");
    const countSpan = document.getElementById("selection-count");
    if (!container) return;
    container.innerHTML = "";
    if (countSpan) countSpan.textContent = `${selectedRecipients.length} 位`;

    selectedRecipients.forEach((user, index) => {
        const tag = document.createElement("span");
        tag.className = "recipient-tag";
        tag.style.display = "inline-block";
        tag.style.background = "#e1f5fe";
        tag.style.padding = "2px 6px";
        tag.style.margin = "2px";
        tag.style.borderRadius = "4px";
        tag.style.fontSize = "0.9em";
        tag.textContent = user.displayName;
        
        const removeBtn = document.createElement("span");
        removeBtn.textContent = " ×";
        removeBtn.style.cursor = "pointer";
        removeBtn.style.color = "red";
        removeBtn.onclick = (e) => {
            e.stopPropagation();
            selectedRecipients.splice(index, 1);
            renderSelectionList();
        };
        
        tag.appendChild(removeBtn);
        container.appendChild(tag);
    });
}

function setupEventHandlers() {
    const clearBtn = document.getElementById("clear-selection-btn");
    if (clearBtn) {
        clearBtn.onclick = () => {
            selectedRecipients = [];
            renderSelectionList();
        };
    }
    document.getElementById("btn-add-to")?.addEventListener("click", () => addRecipientsToOutlook("to"));
    document.getElementById("btn-add-cc")?.addEventListener("click", () => addRecipientsToOutlook("cc"));
    document.getElementById("btn-add-bcc")?.addEventListener("click", () => addRecipientsToOutlook("bcc"));
}

function addRecipientsToOutlook(type) {
    if (selectedRecipients.length === 0) return;
    const recipients = selectedRecipients.map(u => ({
        displayName: u.displayName,
        emailAddress: u.mail || u.userPrincipalName
    }));
    if (Office.context.mailbox.item) {
        Office.context.mailbox.item[type].addAsync(recipients, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) console.error("加入收件人失敗:", result.error);
        });
    }
}