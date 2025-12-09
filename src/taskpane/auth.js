import { PublicClientApplication } from "@azure/msal-browser";

// ⚠️ 這裡必須複製跟 msalClient.js 一模一樣的 Config，除了 redirectUri 以外
// 為了避免循環引用問題，這裡我們直接定義一份乾淨的 Config
const msalConfig = {
  auth: {
    clientId: "11cc40ea-7116-4f77-ae4f-fca0eefbbe4c", // 你的 Client ID
    authority: "https://login.microsoftonline.com/00801dcd-bc88-4134-ad1c-06ebe9f335a6", // 你的 Tenant ID
    redirectUri: window.location.href, // 讓它自動對應當前的 auth.html
  },
  cache: {
    cacheLocation: "localStorage", 
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

(async () => {
  try {
    await msalInstance.initialize();
    // 處理重新導向回來的 Token
    await msalInstance.handleRedirectPromise();
  } catch (error) {
    console.error("Auth Redirect Error:", error);
  }
})();