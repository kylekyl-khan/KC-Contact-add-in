import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

// tenantId / clientId 用你自己的（保持現在能跑的值）
const tenantId = "00801dcd-bc88-4134-ad1c-06ebe9f335a6";
const clientId = "11cc40ea-7116-4f77-ae4f-fca0eefbbe4c";

// ✅ 依照環境，自動決定 redirectUri
function getRedirectUri() {
  // 開發時在 localhost
  if (window.location.hostname === "localhost") {
    return "https://localhost:3000/taskpane.html";
  }
  // 其它情況（例如 Azure），用現在的 origin
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

// Graph 要用到的權限（記得 Entra API permissions 要有這些並且 Admin consent 過）
const loginRequest = {
  scopes: [
    "User.Read",
    "User.Read.All", // 或 Directory.Read.All，看你怎麼設
  ],
};

// 建立 MSAL instance
const msalInstance = new PublicClientApplication(msalConfig);

// 🔑 新版 msal-browser 需要先 initialize
const msalInitPromise = msalInstance.initialize();

/**
 * 確保 MSAL 初始化完成
 */
async function ensureMsalInitialized() {
  await msalInitPromise;
}

/**
 * 確保使用者已登入，沒有登入就跳出登入視窗
 */
export async function ensureLogin() {
  await ensureMsalInitialized();

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    return accounts[0];
  }

  const loginResult = await msalInstance.loginPopup(loginRequest);
  return loginResult.account;
}

/**
 * 取得呼叫 Microsoft Graph 用的 access token
 */
export async function getGraphToken() {
  const account = await ensureLogin();

  const request = {
    ...loginRequest,
    account,
  };

  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const result = await msalInstance.acquireTokenPopup(request);
      return result.accessToken;
    }
    throw e;
  }
}

/**
 * 登出（如果之後要做切換帳號可以用）
 */
export async function logout() {
  await ensureMsalInitialized();

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return;

  await msalInstance.logoutPopup({
    account: accounts[0],
    postLogoutRedirectUri: getRedirectUri(), // 這裡要呼叫函式或用 msalConfig.auth.redirectUri
  });
}