import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

const tenantId = "00801dcd-bc88-4134-ad1c-06ebe9f335a6";
const clientId = "11cc40ea-7116-4f77-ae4f-fca0eefbbe4c";

function getRedirectUri() {
  if (window.location.hostname === "localhost") {
    return "https://localhost:3000/taskpane.html";
  }
  return window.location.origin + "/taskpane.html";
}

const msalConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: getRedirectUri(),
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

const loginRequest = {
  scopes: [
    "User.Read",
    "User.Read.All",
    "Group.Read.All" 
  ],
};

const msalInstance = new PublicClientApplication(msalConfig);

// 初始化狀態旗標
let isInitialized = false;
let initPromise = null;

async function ensureMsalInitialized() {
  if (isInitialized) return;
  if (!initPromise) {
    initPromise = (async () => {
      try {
        await msalInstance.initialize();
        // 處理重導向回來的狀態 (清理 interaction_in_progress)
        await msalInstance.handleRedirectPromise(); 
        isInitialized = true;
      } catch (e) {
        console.error("MSAL Init Error:", e);
        initPromise = null;
        throw e;
      }
    })();
  }
  await initPromise;
}

// 🔥 修改 1: 單純的登入動作 (給 UI 按鈕呼叫用)
export async function loginPopup() {
  await ensureMsalInitialized();
  try {
    const result = await msalInstance.loginPopup(loginRequest);
    return result.account;
  } catch (error) {
    console.error("Login Popup Failed:", error);
    throw error;
  }
}

// 🔥 修改 2: 只嘗試「靜默」獲取 Token，失敗就拋出錯誤，絕不自動彈窗
export async function getGraphToken() {
  await ensureMsalInitialized();
  
  // 檢查是否有帳號資訊
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    // 沒帳號，直接拋出錯誤，讓 UI 顯示登入按鈕
    throw new InteractionRequiredAuthError("No account found");
  }

  const request = { ...loginRequest, account: accounts[0] };

  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (e) {
    console.warn("Silent token acquisition failed:", e);
    // 任何失敗都拋出去，交給 UI 處理
    throw e;
  }
}