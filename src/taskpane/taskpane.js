/* global Office, document */
import { fetchEntraUsers, fetchEntraGroups, fetchGroupMembers } from "../services/graph/entraUsers";
import { loginPopup } from "../services/auth/msalClient"; // 引入登入函式

let allUsers = []; 
let allGroups = [];
let orgTree = null;
let selectedRecipients = [];

const CAMPUS_PREFIX_MAP = {
  "KCQS": "青山校區",
  "KCXG": "秀岡校區",
  "KCHC": "新竹校區",
  "KS": "康軒集團",
  "K1": "康軒集團"
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
    
    // 1. 嘗試抓取群組 (這會觸發 getGraphToken)
    try {
      allGroups = await fetchEntraGroups(); 
      console.log(`✅ 成功抓取群組: ${allGroups.length} 筆`);
      
      // 若成功，繼續正常流程
      loadRestOfApp();

    } catch (e) {
      // 🔥 關鍵修改：如果是驗證錯誤，顯示登入按鈕
      if (e.name === "InteractionRequiredAuthError" || e.message.includes("未登入")) {
          console.log("需要使用者登入，顯示登入按鈕");
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

// 載入應用程式其餘部分 (登入成功後呼叫)
async function loadRestOfApp() {
    // 隱藏登入按鈕 (如果有)
    const loginContainer = document.getElementById("login-container");
    if(loginContainer) loginContainer.remove();

    // 建立樹狀骨架
    orgTree = buildOrgTreeStructure(allGroups);
    renderOrgTree(orgTree); 
    setupEventHandlers();

    // 背景載入使用者
    try {
        const users = await fetchEntraUsers();
        allUsers = users;
    } catch (e) {
        console.warn("⚠️ 無法載入全域使用者:", e);
    }
}

// 🔥 顯示登入按鈕的 UI
function showLoginButton() {
    const appBody = document.getElementById("app-body");
    // 清空內容或覆蓋
    appBody.innerHTML = "";
    
    const container = document.createElement("div");
    container.id = "login-container";
    container.style.display = "flex";
    container.style.flexDirection = "column";
    container.style.alignItems = "center";
    container.style.justifyContent = "center";
    container.style.height = "100%";
    container.style.padding = "20px";
    container.style.textAlign = "center";

    const msg = document.createElement("p");
    msg.textContent = "歡迎使用康橋通訊錄，請先登入以存取資料。";
    msg.style.marginBottom = "20px";
    msg.style.fontSize = "16px";

    const btn = document.createElement("button");
    btn.textContent = "登入 Microsoft 365";
    btn.style.padding = "10px 20px";
    btn.style.fontSize = "16px";
    btn.style.backgroundColor = "#0078d4";
    btn.style.color = "white";
    btn.style.border = "none";
    btn.style.borderRadius = "4px";
    btn.style.cursor = "pointer";

    // 綁定點擊事件 -> 觸發 Popup
    btn.onclick = async () => {
        try {
            await loginPopup(); // 這是使用者主動點擊，瀏覽器不會擋
            // 登入成功後，重新初始化
            // 為了乾淨，簡單地重新整理頁面，或者重新呼叫 initializeOrgUI
            window.location.reload(); 
        } catch (err) {
            console.error("登入失敗:", err);
            msg.textContent = "登入失敗，請重試。";
            msg.style.color = "red";
        }
    };

    container.appendChild(msg);
    container.appendChild(btn);
    appBody.appendChild(container);
}

function showError(text) {
    const appBody = document.getElementById("app-body");
    if(appBody) appBody.innerHTML = `<div style="color:red; padding:20px;">錯誤: ${text}</div>`;
}

// ... (以下 buildOrgTreeStructure, performSearch, createTreeNodeElement, handleAddGroup, renderOrgTree 等函式保持不變，直接貼上您原本的邏輯即可) ...
// === 為了節省篇幅，請保留您原本的這些函式，它們不需要修改 ===

function buildOrgTreeStructure(groups) {
  const root = { id: "root", name: "康橋通訊錄", children: [], users: [] };
  const campusNodes = {};
  const campusPrefixes = Object.keys(CAMPUS_PREFIX_MAP);

  for (const [prefix, name] of Object.entries(CAMPUS_PREFIX_MAP)) {
    if (!campusNodes[name]) {
      const node = { id: `campus-${prefix}`, name: name, children: [], users: [], type: 'campus', membersLoaded: true };
      campusNodes[name] = node;
      root.children.push(node);
    }
  }
  const allNodes = [];
  groups.forEach(g => {
    const match = g.displayName && g.displayName.match(/^([A-Z0-9]+)[\.\-_\s]+(.+)$/);
    if (match) {
      const fullCode = match[1]; 
      const showName = match[2].trim(); 
      let belongingPrefix = null;
      let maxPrefixLen = 0;
      campusPrefixes.forEach(cp => {
          if (fullCode.startsWith(cp) && cp.length > maxPrefixLen) { belongingPrefix = cp; maxPrefixLen = cp.length; }
      });
      if (fullCode === belongingPrefix) return;
      const node = {
          id: g.id, code: fullCode, name: showName, children: [], users: [],
          original: g, membersLoaded: false, isLoading: false, campusPrefix: belongingPrefix
      };
      allNodes.push(node);
    }
  });
  allNodes.sort((a, b) => a.code.length - b.code.length || a.code.localeCompare(b.code));
  allNodes.forEach(childNode => {
      let bestParent = null;
      for (const potentialParent of allNodes) {
          if (potentialParent === childNode) continue;
          if (childNode.code.startsWith(potentialParent.code) && potentialParent.code.length < childNode.code.length) {
              if (!bestParent || potentialParent.code.length > bestParent.code.length) { bestParent = potentialParent; }
          }
      }
      if (bestParent) { bestParent.children.push(childNode); } 
      else {
          const campusName = CAMPUS_PREFIX_MAP[childNode.campusPrefix];
          if (campusName && campusNodes[campusName]) { campusNodes[campusName].children.push(childNode); }
      }
  });
  const codeSort = (a, b) => a.code.localeCompare(b.code);
  Object.values(campusNodes).forEach(c => c.children.sort(codeSort));
  allNodes.forEach(n => { if (n.children.length > 0) n.children.sort(codeSort); });
  return root;
}

function performSearch(keyword) {
    const treeContainer = document.getElementById("org-tree");
    if (!keyword) { renderOrgTree(orgTree); return; }
    treeContainer.innerHTML = "";
    const lowerKey = keyword.toLowerCase();
    const matchedGroups = allGroups.filter(g => g.displayName.toLowerCase().includes(lowerKey));
    const matchedUsers = allUsers.filter(u => u.displayName.toLowerCase().includes(lowerKey) || (u.mail && u.mail.toLowerCase().includes(lowerKey)));
    if (matchedGroups.length === 0 && matchedUsers.length === 0) {
        treeContainer.innerHTML = "<div style='padding:10px; color:#666;'>找不到相符結果</div>";
        return;
    }
    if (matchedGroups.length > 0) {
        const groupHeader = document.createElement("div");
        groupHeader.innerHTML = "<b>📂 相關群組 / 組織</b>";
        groupHeader.style.padding = "5px 10px";
        groupHeader.style.backgroundColor = "#eee";
        treeContainer.appendChild(groupHeader);
        matchedGroups.forEach(g => {
            const mockNode = { id: g.id, name: g.displayName, original: g, children: [], users: [], membersLoaded: false };
            treeContainer.appendChild(createTreeNodeElement(mockNode));
        });
    }
    if (matchedUsers.length > 0) {
        const userHeader = document.createElement("div");
        userHeader.innerHTML = "<b>👤 相關人員</b>";
        userHeader.style.padding = "5px 10px";
        userHeader.style.backgroundColor = "#eee";
        userHeader.style.marginTop = "10px";
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
    row.style.display = "flex";
    row.style.alignItems = "center";
    row.style.justifyContent = "space-between";
    row.style.paddingRight = "10px";
    const titleRow = document.createElement("div");
    titleRow.className = "node-title";
    titleRow.style.cursor = "pointer";
    titleRow.style.padding = "4px";
    titleRow.style.display = "flex";
    titleRow.style.alignItems = "center";
    titleRow.style.flexGrow = "1"; 
    const icon = document.createElement("span");
    const hasChildren = node.children && node.children.length > 0;
    icon.textContent = hasChildren ? "📁 " : "🔹 ";
    icon.style.marginRight = "5px";
    const nameSpan = document.createElement("span");
    nameSpan.textContent = node.name; 
    if (!node.membersLoaded && node.original) { nameSpan.style.color = "#555"; }
    titleRow.appendChild(icon);
    titleRow.appendChild(nameSpan);
    const actionArea = document.createElement("div");
    if (node.original) { 
        const addGroupBtn = document.createElement("span");
        addGroupBtn.textContent = node.original.mail ? " 📧" : " ➕"; 
        addGroupBtn.title = node.original.mail ? `將群組信箱 ${node.original.mail} 加入收件人` : "將群組內所有成員加入收件人";
        addGroupBtn.style.cursor = "pointer";
        addGroupBtn.style.marginLeft = "8px";
        addGroupBtn.style.fontSize = "16px";
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
        icon.textContent = isHidden ? "📂 " : "📁 ";
      }
      if (node.original && !node.membersLoaded) {
          node.isLoading = true;
          nameSpan.textContent = `${node.name} (載入中...)`;
          nameSpan.style.color = "blue";
          try {
              const members = await fetchGroupMembers(node.original.id);
              node.users = members;
              node.membersLoaded = true;
              nameSpan.textContent = `${node.name} (${members.length})`;
              nameSpan.style.color = members.length > 0 ? "black" : "#888";
              nameSpan.style.fontWeight = members.length > 0 ? "bold" : "normal";
          } catch (err) {
              console.error("載入成員失敗:", err);
              nameSpan.textContent = `${node.name} (載入失敗)`;
              nameSpan.style.color = "red";
          } finally { node.isLoading = false; }
      }
      showContacts(node); 
    };
    let childrenContainer = null;
    if (hasChildren) {
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
  if (rootNode && rootNode.children) { rootNode.children.forEach(campus => { treeContainer.appendChild(createTreeNodeElement(campus)); }); }
}

async function handleAddGroup(node) {
    const group = node.original;
    if (group.mail) { addToSelection({ id: group.id, displayName: `[群組] ${group.displayName}`, mail: group.mail, type: 'group' }); return; }
    let members = node.users;
    if (!node.membersLoaded) {
        const btn = document.activeElement;
        if(btn) btn.style.cursor = "wait";
        try { members = await fetchGroupMembers(group.id); node.users = members; node.membersLoaded = true; } 
        catch (e) { console.error("加入群組全員失敗:", e); return; } 
        finally { if(btn) btn.style.cursor = "pointer"; }
    }
    if (members.length === 0) { console.log("群組內無成員"); return; }
    members.forEach(user => addToSelection(user));
}

function createContactItem(user) {
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
    return item;
}

function showContacts(node) {
  const listContainer = document.getElementById("contacts-list");
  if (!listContainer) return;
  listContainer.innerHTML = ""; 
  const breadcrumb = document.getElementById("breadcrumb");
  if (breadcrumb) breadcrumb.textContent = node.name;
  const countSpan = document.getElementById("contacts-count");
  if (countSpan) {
      if (node.membersLoaded) { countSpan.textContent = `共 ${node.users.length} 筆`; } 
      else { countSpan.textContent = "點擊載入..."; }
  }
  if (!node.users || node.users.length === 0) {
    const emptyMsg = document.createElement("div");
    emptyMsg.textContent = node.membersLoaded ? "此群組無成員" : "請點擊群組標題以載入成員";
    emptyMsg.style.color = "#888";
    emptyMsg.style.padding = "10px";
    listContainer.appendChild(emptyMsg);
    return;
  }
  node.users.forEach(user => { listContainer.appendChild(createContactItem(user)); });
}

function addToSelection(userOrGroup) {
    if (selectedRecipients.find(u => u.id === userOrGroup.id)) return;
    selectedRecipients.push(userOrGroup);
    renderSelectionList();
}

function renderSelectionList() {
    const container = document.getElementById("selection-list");
    const countSpan = document.getElementById("selection-count");
    if (!container) return;
    container.innerHTML = "";
    if (countSpan) countSpan.textContent = `${selectedRecipients.length} 位`;
    selectedRecipients.forEach((item, index) => {
        const tag = document.createElement("span");
        tag.className = "recipient-tag";
        tag.style.display = "inline-flex";
        const isGroup = item.type === 'group';
        tag.style.background = isGroup ? "#e0f7fa" : "#deecf9";
        if (isGroup) tag.style.border = "1px solid #006064";
        tag.style.padding = "2px 6px";
        tag.style.margin = "2px";
        tag.style.borderRadius = "4px";
        tag.style.fontSize = "0.9em";
        tag.textContent = item.displayName;
        const removeBtn = document.createElement("span");
        removeBtn.textContent = " ×";
        removeBtn.style.cursor = "pointer";
        removeBtn.style.color = "red";
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

function addRecipientsToOutlook(type) {
    if (selectedRecipients.length === 0) return;
    const recipients = selectedRecipients.map(u => ({ displayName: u.displayName, emailAddress: u.mail || u.userPrincipalName }));
    if (Office.context.mailbox.item) {
        Office.context.mailbox.item[type].addAsync(recipients, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) console.error("加入收件人失敗:", result.error);
        });
    }
}