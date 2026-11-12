# 出金中心

Prop Firm（Lucid / Tradeify）出金與成本統計網頁。

- Google 登入，資料存在 Firebase Firestore（trade-sync-e41ce），每個帳號只能讀寫自己的資料
- 匯入 Lucid Order History / Payout History、Tradeify Orders 從網頁複製到 Excel 的檔案
- FAILED 訂單不記入；以訂單編號自動排除重複
- 可加到手機主畫面當 App 使用

`firestore.rules` 是 Firestore 安全規則，需貼到 Firebase 後台發布。
