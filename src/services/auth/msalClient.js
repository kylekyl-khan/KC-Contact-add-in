const { PublicClientApplication, InteractionRequiredAuthError } = require("@azure/msal-browser");

// TODO: 以下設定請填入實際的 Entra App Registration 參數
// - tenantId: 從 Azure Portal > Microsoft Entra ID > Overview 取得租戶 ID
// - clientId: 從 App registrations > (您的應用程式) > Application (client) ID 取得
// - redirectUri: 本地開發可使用 https://localhost:3000/taskpane.html，正式環境請改為佈署網址
const msalConfig = {
  auth: {
    clientId: "YOUR_CLIENT_ID", // TODO: 代入實際 clientId
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID", // TODO: 代入實際 tenantId
    redirectUri: "https://localhost:3000/taskpane.html", // TODO: 依環境調整 redirectUri
  },
  cache: {
    cacheLocation: "localStorage",
  },
};

const loginRequest = {
  scopes: ["User.Read", "Group.Read.All", "User.ReadBasic.All"],
};

const msalInstance = new PublicClientApplication(msalConfig);

async function ensureLogin() {
  // 檢查是否已有登入帳號，避免重複彈窗
  const accounts = msalInstance.getAllAccounts();
  if (accounts && accounts.length > 0) {
    const activeAccount = accounts[0];
    msalInstance.setActiveAccount(activeAccount);
    return activeAccount;
  }

  const loginResult = await msalInstance.loginPopup(loginRequest);
  if (loginResult && loginResult.account) {
    msalInstance.setActiveAccount(loginResult.account);
    return loginResult.account;
  }
  throw new Error("MSAL login failed: no account returned");
}

async function getGraphToken() {
  const account = await ensureLogin();
  const request = { ...loginRequest, account };

  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const interactiveResult = await msalInstance.acquireTokenPopup(request);
      return interactiveResult.accessToken;
    }
    throw error;
  }
}

async function logout() {
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  await msalInstance.logoutPopup({ account });
}

module.exports = {
  ensureLogin,
  getGraphToken,
  logout,
  msalInstance,
};
