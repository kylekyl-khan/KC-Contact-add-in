# Entra 群組與 Task Pane 假資料流程

此文件說明目前 Task Pane 如何使用 Entra 匯出的群組 CSV 產生示範用的通訊錄資料，以及未來切換到 Microsoft Graph 的位置。

## 檔案結構
- `data/exportGroups_2025-11-19.csv`：從 Entra 匯出的群組清單（含 id / displayName 等欄位）。
- `src/data/orgTreeConfig.js`：白名單組織樹設定，只會在 UI 中露出這些群組。
- `scripts/generateOrgMembersFromCsv.js`：讀取 CSV 並針對 orgTreeConfig 中的 groupId 產生假成員。
- `data/orgMembers.generated.json`：由腳本產生的成員快取，Task Pane 直接 import。
- `src/data/entraMembersService.js`：封裝讀取 `orgMembers.generated.json` 的介面，未來可以改成呼叫 Microsoft Graph。

## 重新產生 orgMembers.generated.json
1. 更新或覆蓋 `data/exportGroups_2025-11-19.csv`（確保有最新的群組 id）。
2. 編輯 `src/data/orgTreeConfig.js` 確認需要顯示的 groupId 已列入白名單。
3. 執行：
   ```bash
   node scripts/generateOrgMembersFromCsv.js
   ```
   會在 `data/orgMembers.generated.json` 產生示範成員。當前腳本用群組名稱組合出三筆假資料，真正接 Graph 時可替換腳本邏輯。

## 開發與預覽
- 啟動 dev server：
  ```bash
  npm run dev-server
  ```
- 在瀏覽器開啟 `https://localhost:3000/taskpane.html` 即可預覽「預覽模式」，僅更新 Task Pane 下方的選擇清單。
- 若在 Outlook Add-in 環境（Office.onReady 成功、host=Outlook），點「加到收件人」會呼叫 `Office.context.mailbox.item.to.addAsync` 寫入 To 欄位。

## 未來接 Microsoft Graph 的建議
- `src/data/entraMembersService.js` 的 `getMembersByGroupId` 目前直接讀取 `orgMembers.generated.json`。
- 未來可在該 function 中改為使用 `@microsoft/microsoft-graph-client`：
  1. 透過 Azure AD 驗證取得 access token。
  2. 呼叫 `/groups/{groupId}/members` 或 `/groups/{groupId}/transitiveMembers` 取得實際人員。
  3. 回傳格式保持 `{ id, name, title, email }`，以便 Task Pane 無須修改其他程式碼。
