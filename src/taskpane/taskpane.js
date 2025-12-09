/* global Office, document */
import { fetchEntraUsers, fetchEntraGroups, fetchGroupMembers } from "../services/graph/entraUsers";
import { loginPopup } from "../services/auth/msalClient";

let allUsers = []; 
let allGroups = [];
let orgTree = null;
let selectedRecipients = [];

// 🔥 設定：校區對照表
const CAMPUS_PREFIX_MAP = {
  "KCQS": "青山校區",
  "KCXG": "秀岡校區",
  "KCHC": "新竹校區"
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
    console.log("🚀 開始初始化...");
    
    try {
      allGroups = await fetchEntraGroups(); 
      console.log(`✅ 成功抓取群組: ${allGroups.length} 筆`);
      loadRestOfApp();
    } catch (e) {
      if (e.name === "InteractionRequiredAuthError" || e.message.includes("未登入")) {
          showLoginButton();
      } else {
          console.error("其他錯誤:", e);
          showError(`系統錯誤: ${e.message}`);
      }
    }
  } catch (e) {
    console.error("💥 初始化錯誤：", e);
    showError(e.message);
  }
}

async function loadRestOfApp() {
    const loginContainer = document.getElementById("login-container");
    if(loginContainer) loginContainer.remove();

    orgTree = buildOrgTreeStructure(allGroups);
    renderOrgTree(orgTree); 
    setupEventHandlers();

    // 預載使用者 (選用)
    try {
        const users = await fetchEntraUsers();
        allUsers = users;
    } catch (e) {
        console.warn("⚠️ 無法載入全域使用者:", e);
    }
}

// ---------------------------------------------------------------------------
// 🔥 核心邏輯：建立樹狀結構 (放寬篩選版)
// ---------------------------------------------------------------------------
function buildOrgTreeStructure(groups) {
  console.log("🌳 開始建立組織樹 (寬鬆版)...");
  const root = { id: "root", name: "康橋通訊錄", children: [], users: [] };
  const campusNodes = {};

  // 1. 初始化三個校區的根節點
  for (const [prefix, name] of Object.entries(CAMPUS_PREFIX_MAP)) {
    const node = { 
        id: `campus-${prefix}`, 
        name: name, 
        children: [], 
        users: [], 
        type: 'campus', 
        membersLoaded: true 
    };
    campusNodes[name] = node;
    campusNodes[prefix] = node; 
    root.children.push(node);
  }

  const validNodes = [];

  groups.forEach(g => {
    let fullCode = "";
    let showName = g.displayName;
    let belongingPrefix = null;

    // 步驟 A: 嘗試用 Regex 解析標準格式 (代碼 - 名稱)
    const match = g.displayName && g.displayName.match(/^([A-Z0-9]+)[\.\-_\s]+(.+)$/);
    
    if (match) {
        fullCode = match[1]; 
        showName = match[2].trim();
        
        // 檢查代碼是否符合校區前綴
        let maxPrefixLen = 0;
        Object.keys(CAMPUS_PREFIX_MAP).forEach(cp => {
            if (fullCode.startsWith(cp) && cp.length > maxPrefixLen) { 
                belongingPrefix = cp; 
                maxPrefixLen = cp.length; 
            }
        });
    }

    // 步驟 B: 如果 Regex 沒抓到，改用關鍵字搜尋 (放寬條件)
    if (!belongingPrefix) {
        // 檢查名稱是否包含中文校區名 (例如 "秀岡")
        for (const [prefix, name] of Object.entries(CAMPUS_PREFIX_MAP)) {
            // 去掉"校區"兩個字來比對，增加命中率 (ex: "秀岡教務處" 也能對應 "秀岡校區")
            const shortName = name.replace("校區", ""); 
            if (g.displayName.includes(shortName) || g.displayName.startsWith(prefix)) {
                belongingPrefix = prefix;
                // 如果沒有代碼，就用整個名稱當顯示名稱
                fullCode = ""; 
                showName = g.displayName;
                break;
            }
        }
    }

    // ⚠️ 最終過濾：如果還是找不到歸屬，就真的跳過
    if (!belongingPrefix) return;

    const node = {
        id: g.id, 
        code: fullCode, 
        name: showName, 
        displayName: g.displayName, 
        children: [], 
        users: [],
        original: g, 
        membersLoaded: false,
        isLoading: false, 
        campusPrefix: belongingPrefix
    };
    validNodes.push(node);
  });

  console.log(`🌲 篩選後保留節點: ${validNodes.length} / ${groups.length}`);

  // 2. 自動層級組裝 (如果有代碼的話)
  validNodes.sort((a, b) => {
      const codeA = a.code || "";
      const codeB = b.code || "";
      return codeA.length - codeB.length || codeA.localeCompare(codeB);
  });
  
  validNodes.forEach(childNode => {
      let bestParent = null;

      // 只有當此節點有代碼時，才嘗試尋找父節點
      if (childNode.code) {
          for (const potentialParent of validNodes) {
              if (potentialParent === childNode) continue;
              if (!potentialParent.code) continue; // 父節點也必須有代碼
              
              if (childNode.code.startsWith(potentialParent.code) && potentialParent.code.length < childNode.code.length) {
                  if (!bestParent || potentialParent.code.length > bestParent.code.length) { 
                      bestParent = potentialParent; 
                  }
              }
          }
      }
      
      if (bestParent) { 
          bestParent.children.push(childNode); 
      } else {
          // 沒爸爸，加入校區根目錄
          const campusNode = campusNodes[childNode.campusPrefix];
          if (campusNode) { 
              campusNode.children.push(childNode);
          }
      }
  });

  // 3. 排序顯示
  const recursiveSort = (nodes) => {
      nodes.sort((a, b) => {
          const codeA = a.code || "";
          const codeB = b.code || "";
          // 有代碼的排前面，沒代碼的照名稱排
          if(codeA && !codeB) return -1;
          if(!codeA && codeB) return 1;
          if(!codeA && !codeB) return a.name.localeCompare(b.name);
          return codeA.localeCompare(codeB);
      });
      nodes.forEach(n => {
          if (n.children && n.children.length > 0) {
              recursiveSort(n.children);
          }
      });
  };

  root.children.forEach(campus => {
      if (campus.children.length > 0) {
          recursiveSort(campus.children);
      }
  });

  return root;
}

function performSearch(keyword) {
    const treeContainer = document.getElementById("org-tree");
    if (!keyword) { renderOrgTree(orgTree); return; }
    treeContainer.innerHTML = "";
    
    const lowerKey = keyword.toLowerCase();
    const matchedGroups = allGroups.filter(g => {
        const isTargetCampus = Object.keys(CAMPUS_PREFIX_MAP).some(prefix => 
            g.displayName.startsWith(prefix) || 
            g.displayName.includes(CAMPUS_PREFIX_MAP[prefix].replace("校區", ""))
        );
        return isTargetCampus && g.displayName.toLowerCase().includes(lowerKey);
    });
    
    const matchedUsers = allUsers.filter(u => u.displayName.toLowerCase().includes(lowerKey) || (u.mail && u.mail.toLowerCase().includes(lowerKey)));
    
    if (matchedGroups.length === 0 && matchedUsers.length === 0) {
        treeContainer.innerHTML = "<div style='padding:10px; color:#666;'>找不到相符結果</div>";
        return;
    }

    if (matchedGroups.length > 0) {
        const groupHeader = document.createElement("div");
        groupHeader.innerHTML = "<b>📂 相關群組</b>";
        groupHeader.style.cssText = "padding:5px 10px; background:#eee; margin-bottom:5px;";
        treeContainer.appendChild(groupHeader);
        matchedGroups.forEach(g => {
            const mockNode = { id: g.id, name: g.displayName, original: g, children: [], users: [], membersLoaded: false };
            treeContainer.appendChild(createTreeNodeElement(mockNode));
        });
    }

    if (matchedUsers.length > 0) {
        const userHeader = document.createElement("div");
        userHeader.innerHTML = "<b>👤 相關人員</b>";
        userHeader.style.cssText = "padding:5px 10px; background:#eee; margin-top:10px; margin-bottom:5px;";
        treeContainer.appendChild(userHeader);
        const listDiv = document.createElement("div");
        matchedUsers.forEach(user => { listDiv.appendChild(createContactItem(user)); });
        treeContainer.appendChild(listDiv);
    }
}

function createTreeNodeElement(node) {
    const nodeEl = document.createElement("div");
    nodeEl.className = "tree-node";
    nodeEl.style.marginLeft = "15px";
    
    const row = document.createElement("div");
    row.style.cssText = "display:flex; align-items:center; justify-content:space-between; padding-right:10px;";
    
    const titleRow = document.createElement("div");
    titleRow.className = "node-title";
    titleRow.style.cssText = "cursor:pointer; padding:6px; display:flex; align-items:center; flex-grow:1; border-radius:4px;";
    titleRow.onmouseover = () => titleRow.style.backgroundColor = "#f0f0f0";
    titleRow.onmouseout = () => titleRow.style.backgroundColor = "transparent";

    const icon = document.createElement("span");
    const isFolder = (node.children && node.children.length > 0) || node.type === 'campus';
    icon.textContent = isFolder ? "📁 " : "🔹 ";
    icon.style.marginRight = "6px";
    
    const nameSpan = document.createElement("span");
    nameSpan.textContent = node.name;
    if (!node.membersLoaded && node.original) { nameSpan.style.color = "#555"; }
    
    titleRow.appendChild(icon);
    titleRow.appendChild(nameSpan);
    
    const actionArea = document.createElement("div");
    if (node.original) { 
        const addGroupBtn = document.createElement("span");
        addGroupBtn.textContent = "➕"; 
        addGroupBtn.title = "將群組成員加入收件人";
        addGroupBtn.style.cssText = "cursor:pointer; margin-left:8px; font-size:14px; padding:2px 6px; border:1px solid #ccc; border-radius:4px;";
        addGroupBtn.onclick = async (e) => { e.stopPropagation(); await handleAddGroup(node); };
        actionArea.appendChild(addGroupBtn);
    }
    
    row.appendChild(titleRow);
    row.appendChild(actionArea);
    nodeEl.appendChild(row);

    titleRow.onclick = async (e) => {
      e.stopPropagation();
      if (node.isLoading) return;

      if (childrenContainer) {
        const isHidden = childrenContainer.style.display === "none";
        childrenContainer.style.display = isHidden ? "block" : "none";
        if(isFolder) icon.textContent = isHidden ? "📂 " : "📁 ";
      }

      if (node.original && !node.membersLoaded) {
          node.isLoading = true;
          nameSpan.textContent = `${node.name} (載入中...)`;
          try {
              const members = await fetchGroupMembers(node.original.id);
              node.users = members;
              node.membersLoaded = true;
              const count = members.length;
              nameSpan.textContent = `${node.name} (${count})`;
              nameSpan.style.fontWeight = count > 0 ? "bold" : "normal";
              nameSpan.style.color = count > 0 ? "black" : "#888";
          } catch (err) {
              console.error("載入失敗:", err);
              nameSpan.textContent = `${node.name} (失敗)`;
              nameSpan.style.color = "red";
          } finally { 
              node.isLoading = false; 
          }
      }
      showContacts(node); 
    };

    let childrenContainer = null;
    if (node.children && node.children.length > 0) {
      childrenContainer = document.createElement("div");
      childrenContainer.className = "node-children";
      childrenContainer.style.display = "none"; 
      node.children.forEach(child => { childrenContainer.appendChild(createTreeNodeElement(child)); });
      nodeEl.appendChild(childrenContainer);
    }
    return nodeEl;
}

function renderOrgTree(rootNode) {
  const treeContainer = document.getElementById("org-tree");
  if (!treeContainer) return;
  treeContainer.innerHTML = ""; 
  if (rootNode && rootNode.children) { 
      rootNode.children.forEach(child => treeContainer.appendChild(createTreeNodeElement(child))); 
  }
}

async function handleAddGroup(node) {
    const group = node.original;
    if (!node.membersLoaded) {
        try { 
            const members = await fetchGroupMembers(group.id); 
            node.users = members; 
            node.membersLoaded = true; 
        } catch (e) { 
            console.error("加入群組失敗:", e); 
            return; 
        } 
    }
    
    if (!node.users || node.users.length === 0) {
        // 🔥 修正: 移除 alert，改用 console 警告
        console.warn("此群組沒有成員，無法加入。");
        return; 
    }
    
    node.users.forEach(user => addToSelection(user));
}

function createContactItem(user) {
    const item = document.createElement("div");
    item.className = "contact-item";
    item.style.cssText = "padding:10px; border-bottom:1px solid #f0f0f0; display:flex; justify-content:space-between; align-items:center; cursor:pointer;";
    item.onmouseover = () => item.style.backgroundColor = "#fafafa";
    item.onmouseout = () => item.style.backgroundColor = "transparent";

    const infoDiv = document.createElement("div");
    const nameDiv = document.createElement("div");
    nameDiv.textContent = user.displayName;
    nameDiv.style.fontWeight = "bold";
    
    const emailDiv = document.createElement("div");
    emailDiv.textContent = user.mail || user.userPrincipalName || "無 Email";
    emailDiv.style.fontSize = "12px";
    emailDiv.style.color = "#666";
    
    if (user.jobTitle) {
        const jobSpan = document.createElement("span");
        jobSpan.textContent = ` • ${user.jobTitle}`;
        jobSpan.style.fontSize = "12px";
        jobSpan.style.color = "#888";
        nameDiv.appendChild(jobSpan);
    }

    infoDiv.appendChild(nameDiv);
    infoDiv.appendChild(emailDiv);

    const addBtn = document.createElement("button");
    addBtn.textContent = "+";
    addBtn.style.cssText = "padding:2px 10px; border:1px solid #ddd; background:white; cursor:pointer; border-radius:4px;";
    
    item.appendChild(infoDiv);
    item.appendChild(addBtn);
    item.onclick = () => addToSelection(user);
    return item;
}

function showContacts(node) {
  const listContainer = document.getElementById("contacts-list");
  if (!listContainer) return;
  listContainer.innerHTML = ""; 
  
  const breadcrumb = document.getElementById("breadcrumb");
  if (breadcrumb) breadcrumb.textContent = node.name || "群組成員";
  
  const countSpan = document.getElementById("contacts-count");
  if (countSpan) countSpan.textContent = node.membersLoaded ? `共 ${node.users.length} 筆` : "";

  if (!node.users || node.users.length === 0) {
    const emptyMsg = document.createElement("div");
    emptyMsg.style.padding = "20px";
    emptyMsg.style.color = "#666";
    emptyMsg.style.textAlign = "center";
    
    if (node.membersLoaded) {
        emptyMsg.textContent = "此群組無成員";
        const hint = document.createElement("div");
        hint.textContent = "(API 回傳 0 筆資料)";
        hint.style.fontSize = "12px";
        hint.style.marginTop = "5px";
        emptyMsg.appendChild(hint);
    } else {
        emptyMsg.textContent = "👈 請點擊左側群組以載入成員";
    }
    listContainer.appendChild(emptyMsg);
    return;
  }
  
  node.users.forEach(user => { listContainer.appendChild(createContactItem(user)); });
}

function addToSelection(user) {
    if (!user.mail && !user.userPrincipalName) return; 
    if (selectedRecipients.find(u => u.id === user.id)) return;
    selectedRecipients.push(user);
    renderSelectionList();
}

function renderSelectionList() {
    const container = document.getElementById("selection-list");
    const countSpan = document.getElementById("selection-count");
    if (!container) return;
    container.innerHTML = "";
    if (countSpan) countSpan.textContent = selectedRecipients.length;
    
    selectedRecipients.forEach((item, index) => {
        const tag = document.createElement("span");
        tag.className = "recipient-tag";
        tag.textContent = item.displayName;
        
        const removeBtn = document.createElement("span");
        removeBtn.textContent = " ×";
        removeBtn.onclick = (e) => { e.stopPropagation(); selectedRecipients.splice(index, 1); renderSelectionList(); };
        tag.appendChild(removeBtn);
        container.appendChild(tag);
    });
}

function setupEventHandlers() {
    const searchInput = document.getElementById("search-input");
    const clearSearchBtn = document.getElementById("clear-search-btn");
    if (searchInput) { searchInput.addEventListener("input", (e) => { performSearch(e.target.value.trim()); }); }
    if (clearSearchBtn) { clearSearchBtn.addEventListener("click", () => { if (searchInput) { searchInput.value = ""; performSearch(""); } }); }
    const clearBtn = document.getElementById("clear-selection-btn");
    if (clearBtn) { clearBtn.onclick = () => { selectedRecipients = []; renderSelectionList(); }; }
    document.getElementById("btn-add-to")?.addEventListener("click", () => addRecipientsToOutlook("to"));
    document.getElementById("btn-add-cc")?.addEventListener("click", () => addRecipientsToOutlook("cc"));
    document.getElementById("btn-add-bcc")?.addEventListener("click", () => addRecipientsToOutlook("bcc"));
}

function showLoginButton() {
    const appBody = document.getElementById("app-body");
    appBody.innerHTML = "";
    const container = document.createElement("div");
    container.id = "login-container";
    container.style.cssText = "display:flex; flex-direction:column; align-items:center; justify-content:center; height:100%; padding:20px; text-align:center;";
    const msg = document.createElement("p");
    msg.textContent = "歡迎使用康橋通訊錄，請先登入以存取資料。";
    msg.style.marginBottom = "20px";
    const btn = document.createElement("button");
    btn.textContent = "登入 Microsoft 365";
    btn.style.cssText = "padding:10px 20px; background-color:#0078d4; color:white; border:none; border-radius:4px; cursor:pointer;";
    btn.onclick = async () => {
        try { await loginPopup(); window.location.reload(); } 
        catch (err) { console.error("登入失敗:", err); msg.textContent = "登入失敗，請重試。"; msg.style.color = "red"; }
    };
    container.appendChild(msg);
    container.appendChild(btn);
    appBody.appendChild(container);
}

function showError(text) {
    const appBody = document.getElementById("app-body");
    if(appBody) appBody.innerHTML = `<div style="color:red; padding:20px;">錯誤: ${text}</div>`;
}

function addRecipientsToOutlook(type) {
    if (selectedRecipients.length === 0) return;
    const recipients = selectedRecipients.map(u => ({ displayName: u.displayName, emailAddress: u.mail || u.userPrincipalName }));
    if (Office.context.mailbox && Office.context.mailbox.item) {
        Office.context.mailbox.item[type].addAsync(recipients, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) console.error("加入收件人失敗:", result.error);
        });
    } else {
        console.warn("目前不在 Outlook 環境中，無法執行加入動作");
    }
}