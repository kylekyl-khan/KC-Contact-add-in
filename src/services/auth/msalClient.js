import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

// tenantId / clientId 維持原本設定
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

// 🔥 重點修改：加入 Group.Read.All
const loginRequest = {
  scopes: [
    "User.Read",
    "User.Read.All",
    "Group.Read.All" 
  ],
};

const msalInstance = new PublicClientApplication(msalConfig);
const msalInitPromise = msalInstance.initialize();

async function ensureMsalInitialized() {
  await msalInitPromise;
}

export async function ensureLogin() {
  await ensureMsalInitialized();
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) return accounts[0];
  const loginResult = await msalInstance.loginPopup(loginRequest);
  return loginResult.account;
}

export async function getGraphToken() {
  const account = await ensureLogin();
  const request = { ...loginRequest, account };
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

export async function logout() {
  await ensureMsalInitialized();
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return;
  await msalInstance.logoutPopup({
    account: accounts[0],
    postLogoutRedirectUri: getRedirectUri(),
  });
}