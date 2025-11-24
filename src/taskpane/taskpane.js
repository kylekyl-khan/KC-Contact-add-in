/* src/taskpane/taskpane.js */

// ---- Entra 群組設定與成員快取 ----
// orgTreeConfig 控制 UI 露出的群組白名單；實際的成員資料目前從 CSV snapshot 轉成 JSON。
const { orgTreeConfig } = require("../data/orgTreeConfig");
const { getAllMembersWithPath, loadMembersForGroup } = require("../data/entraMembersService");
const { ensureLogin, getGraphToken } = require("../services/auth/msalClient");

// 以設定檔為基礎的組織樹（避免直接修改 import 內容）
const orgTree = JSON.parse(JSON.stringify(orgTreeConfig));

// ---- 全域狀態 ----
let appInitialized = false;
let isOutlook = false;          // 之後真的跑在 Outlook 裡會變 true
let selectedRecipients = [];    // 已選擇收件人列表
let lastSelectedNodeId = null;  // 上一次選到的樹葉節點
let currentMode = "browse";     // "browse" or "search"
let lastSearchKeyword = "";

// ---- 啟動入口：同支援 Outlook & 瀏覽器預覽 ----
function initApp() {
  if (appInitialized) return;
  appInitialized = true;

  // 綁定搜尋與清除按鈕
  const searchInput = document.getElementById("search-input");
  const clearSearchBtn = document.getElementById("clear-search-btn");
  const clearSelectionBtn = document.getElementById("clear-selection-btn");
  const previewBadge = document.querySelector(".app-badge");

  if (previewBadge) {
    previewBadge.textContent = isOutlook ? "Outlook 模式" : "預覽模式";
    previewBadge.classList.toggle("is-outlook", isOutlook);
  }

  if (searchInput) {
    searchInput.addEventListener("input", handleSearchInputChange);
  }

  if (clearSearchBtn) {
    clearSearchBtn.addEventListener("click", () => {
      if (!searchInput) return;
      searchInput.value = "";
      lastSearchKeyword = "";
      currentMode = "browse";

      if (lastSelectedNodeId) {
        selectNode(lastSelectedNodeId);
      } else {
        const firstLeaf = findFirstLeaf(orgTree[0]);
        if (firstLeaf) selectNode(firstLeaf.id);
      }
    });
  }

  if (clearSelectionBtn) {
    clearSelectionBtn.addEventListener("click", () => {
      selectedRecipients = [];
      renderSelection();
    });
  }

  renderOrgTree();

  // 預設選第一個葉節點
  const firstLeaf = findFirstLeaf(orgTree[0]);
  if (firstLeaf) {
    selectNode(firstLeaf.id);
  }

  renderSelection();
}

async function initOutlookMode() {
  const previewBadge = document.querySelector(".app-badge");
  if (previewBadge) {
    previewBadge.textContent = "Outlook 模式";
    previewBadge.classList.add("is-outlook");
  }

  try {
    await ensureLogin();
    const token = await getGraphToken();
    if (token) {
      console.log("Graph token acquired:", `${token.substring(0, 12)}...`);
    }
    initApp();
  } catch (error) {
    console.error("初始化 Outlook 模式失敗：", error);
    renderError("登入 Microsoft 失敗，請稍後再試。");
  }
}

// Outlook 環境：用 Office.onReady 啟動
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(info => {
    try {
      if (info && info.host === Office.HostType.Outlook) {
        isOutlook = true;
      }
    } catch (e) {
      console.warn("Office.onReady info error:", e);
    }

    if (isOutlook) {
      // 確保 DOM 已載入再做 MSAL 初始化
      if (document.readyState === "loading") {
        window.addEventListener("DOMContentLoaded", initOutlookMode);
      } else {
        initOutlookMode();
      }
    } else {
      window.addEventListener("DOMContentLoaded", initApp);
    }
  });
} else {
  // 純瀏覽器預覽：用 DOMContentLoaded 啟動
  window.addEventListener("DOMContentLoaded", initApp);
}

// ------- 樹狀選單渲染 --------

function renderOrgTree() {
  const treeRoot = document.getElementById("org-tree");
  if (!treeRoot) return;
  treeRoot.innerHTML = "";
  orgTree.forEach(node => {
    const el = buildTreeNode(node, 0);
    treeRoot.appendChild(el);
  });
}

function buildTreeNode(node, depth) {
  const container = document.createElement("div");
  container.className = "tree-node";
  container.style.paddingLeft = `${depth * 16}px`;

  const hasChildren = Array.isArray(node.children) && node.children.length > 0;
  const isLeaf = !hasChildren && (!!node.groupId || Array.isArray(node.employees));

  const label = document.createElement("div");
  label.className = "tree-label";
  label.textContent = node.name;

  label.addEventListener("click", () => {
    if (isLeaf) {
      selectNode(node.id);
    } else if (hasChildren) {
      // 展開 / 收合
      container.classList.toggle("collapsed");
    }
  });

  container.appendChild(label);

  if (hasChildren) {
    const childrenContainer = document.createElement("div");
    childrenContainer.className = "tree-children";
    node.children.forEach(child => {
      childrenContainer.appendChild(buildTreeNode(child, depth + 1));
    });
    container.appendChild(childrenContainer);
  }

  if (isLeaf) {
    container.dataset.nodeId = node.id;
  }

  return container;
}

// ------- 節點選取 & 顯示員工 --------

async function selectNode(nodeId) {
  lastSelectedNodeId = nodeId;
  currentMode = "browse";
  lastSearchKeyword = "";

  const searchInput = document.getElementById("search-input");
  if (searchInput) searchInput.value = "";

  // 高亮目前選取的節點
  document
    .querySelectorAll(".tree-node.leaf-selected")
    .forEach(el => el.classList.remove("leaf-selected"));

  const leaf = document.querySelector(`.tree-node[data-node-id="${nodeId}"]`);
  if (leaf) {
    leaf.classList.add("leaf-selected");
  }

  const node = findNodeById(orgTree, nodeId);
  if (!node) return;

  // 有 employees 為 demo 節點，直接渲染
  if (Array.isArray(node.employees)) {
    renderEmployees(node.employees, buildBreadcrumb(node), { isSearch: false });
    return;
  }

  if (node.groupId) {
    renderLoading(buildBreadcrumb(node));
    try {
      const employees = await loadMembersForGroup(node.groupId);
      renderEmployees(employees, buildBreadcrumb(node), { isSearch: false });
    } catch (err) {
      console.error(err);
      renderError("載入群組成員時發生錯誤，請稍後再試一次。");
    }
    return;
  }

  renderEmployees([], buildBreadcrumb(node), { isSearch: false });
}

function renderLoading(breadcrumbText) {
  const breadcrumb = document.getElementById("breadcrumb");
  const list = document.getElementById("contacts-list");
  const countEl = document.getElementById("contacts-count");
  if (breadcrumb) breadcrumb.textContent = breadcrumbText || "";
  if (countEl) countEl.textContent = "載入中...";
  if (list) list.innerHTML = "<div class='empty'>載入中...</div>";
}

function renderError(message) {
  const list = document.getElementById("contacts-list");
  const countEl = document.getElementById("contacts-count");
  if (countEl) countEl.textContent = "0";
  if (list) list.innerHTML = `<div class='empty'>${message}</div>`;
}

function findNodeById(nodes, id) {
  for (const n of nodes) {
    if (n.id === id) return n;
    if (n.children) {
      const child = findNodeById(n.children, id);
      if (child) return child;
    }
  }
  return null;
}

function findFirstLeaf(node) {
  if ((node.groupId || (node.employees && node.employees.length > 0)) && !node.children) return node;
  if (!node.children) return null;
  for (const child of node.children) {
    const leaf = findFirstLeaf(child);
    if (leaf) return leaf;
  }
  return null;
}

function buildBreadcrumb(node) {
  if (!node) return "";
  const path = [];
  let current = node;
  while (current) {
    path.unshift(current.name);
    current = findParent(orgTree[0], current.id);
  }
  return path.join(" / ");
}

function findParent(root, childId, parent = null) {
  if (!root) return null;
  if (root.id === childId) return parent;
  if (root.children) {
    for (const c of root.children) {
      const found = findParent(c, childId, root);
      if (found) return found;
    }
  }
  return null;
}

// ------- 搜尋功能：全公司搜尋姓名 / 職稱 / Email --------

async function handleSearchInputChange(event) {
  const keyword = event.target.value.trim();
  lastSearchKeyword = keyword;

  if (!keyword) {
    currentMode = "browse";
    if (lastSelectedNodeId) {
      selectNode(lastSelectedNodeId);
    } else {
      const firstLeaf = findFirstLeaf(orgTree[0]);
      if (firstLeaf) selectNode(firstLeaf.id);
    }
    return;
  }

  currentMode = "search";
  renderLoading(`搜尋結果：「${keyword}」`);
  try {
    const results = await searchEmployees(keyword);
    renderEmployees(results, `搜尋結果：「${keyword}」`, { isSearch: true, showOrgPath: true });
  } catch (err) {
    console.error(err);
    renderError("搜尋時發生錯誤，請稍後再試一次。");
  }
}

async function searchEmployees(keyword) {
  const all = await collectAllEmployeesWithPath(orgTree);
  const lower = keyword.toLowerCase();
  return all.filter(emp =>
    (emp.name && emp.name.toLowerCase().includes(lower)) ||
    (emp.title && emp.title.toLowerCase().includes(lower)) ||
    (emp.email && emp.email.toLowerCase().includes(lower))
  );
}

async function collectAllEmployeesWithPath(nodes) {
  // 先優先使用服務整合的快取，避免重複展開
  const enriched = await getAllMembersWithPath();
  if (enriched.length) return enriched;

  // 若無預載資料，仍 fallback 以目前節點資料搜尋
  const walk = (list, path = []) => {
    let result = [];
    for (const n of list) {
      const newPath = [...path, n.name];
      if (Array.isArray(n.employees)) {
        result = result.concat(
          n.employees.map(e => ({
            ...e,
            path: newPath.join(" / ")
          }))
        );
      }
      if (Array.isArray(n.children)) {
        result = result.concat(walk(n.children, newPath));
      }
    }
    return result;
  };
  return walk(nodes);
}

// ------- 員工卡片 & 清單渲染 --------

function renderEmployees(employees, breadcrumbText, options = {}) {
  const breadcrumb = document.getElementById("breadcrumb");
  const list = document.getElementById("contacts-list");
  const countEl = document.getElementById("contacts-count");

  if (breadcrumb) {
    breadcrumb.textContent = breadcrumbText || "";
  }

  if (!list) return;
  list.innerHTML = "";

  if (countEl) {
    countEl.textContent = `共 ${employees.length} 筆`;
  }

  if (!employees.length) {
    list.innerHTML = "<div class='empty'>沒有符合條件的員工。</div>";
    return;
  }

  employees.forEach(emp => {
    const card = document.createElement("div");
    card.className = "contact-card";

    const nameEl = document.createElement("div");
    nameEl.className = "contact-name";
    nameEl.textContent = emp.name;

    const titleEl = document.createElement("div");
    titleEl.className = "contact-title";
    titleEl.textContent = emp.title || "";

    const emailEl = document.createElement("div");
    emailEl.className = "contact-email";
    emailEl.textContent = emp.email;

    card.appendChild(nameEl);
    card.appendChild(titleEl);

    if (options.showOrgPath && emp.path) {
      const pathEl = document.createElement("div");
      pathEl.className = "contact-path";
      pathEl.textContent = emp.path;
      card.appendChild(pathEl);
    }

    card.appendChild(emailEl);

    const btn = document.createElement("button");
    btn.className = "contact-add-btn";
    btn.textContent = isOutlook ? "加到收件人" : "加入選擇清單";
    btn.addEventListener("click", () => addToRecipient(emp));

    card.appendChild(btn);
    list.appendChild(card);
  });
}

// ------- 已選擇收件人（預覽模式 + 之後可共用） --------

function addToRecipient(emp) {
  // 預覽 / 非 Outlook 環境：只操作本地選擇清單
  if (!isOutlook || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item || !Office.context.mailbox.item.to) {
    addToLocalSelection(emp);
    return;
  }

  try {
    const item = Office.context.mailbox.item;
    const recipient = {
      emailAddress: emp.email,
      displayName: emp.name
    };

    item.to.addAsync([recipient], result => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error(result.error);
        alert("加入收件人失敗：" + result.error.message);
      } else {
        // 也同步到本地選擇清單，方便 Task Pane 顯示
        addToLocalSelection(emp);
      }
    });
  } catch (e) {
    console.error(e);
    alert("Outlook API 尚未準備好，請稍後再試一次。");
  }
}

function addToLocalSelection(emp) {
  if (!emp || !emp.email) return;
  if (!selectedRecipients.some(e => e.email === emp.email)) {
    selectedRecipients.push({
      name: emp.name,
      title: emp.title,
      email: emp.email
    });
    renderSelection();
  }
}

function renderSelection() {
  const list = document.getElementById("selection-list");
  const countEl = document.getElementById("selection-count");
  if (!list) return;

  list.innerHTML = "";

  if (!selectedRecipients.length) {
    list.innerHTML = "<div class='empty-selection'>尚未選擇任何收件人。</div>";
    if (countEl) countEl.textContent = "0 位";
    return;
  }

  selectedRecipients.forEach((emp, index) => {
    const row = document.createElement("div");
    row.className = "selection-row";

    const main = document.createElement("div");
    main.className = "selection-main";
    main.textContent = `${emp.name}${emp.title ? "（" + emp.title + "）" : ""}`;

    const email = document.createElement("div");
    email.className = "selection-email";
    email.textContent = emp.email;

    const removeBtn = document.createElement("button");
    removeBtn.className = "selection-remove-btn";
    removeBtn.textContent = "移除";
    removeBtn.addEventListener("click", () => {
      selectedRecipients.splice(index, 1);
      renderSelection();
    });

    row.appendChild(main);
    row.appendChild(email);
    row.appendChild(removeBtn);

    list.appendChild(row);
  });

  if (countEl) {
    countEl.textContent = `${selectedRecipients.length} 位`;
  }
}
