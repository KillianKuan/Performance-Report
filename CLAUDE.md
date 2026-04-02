
---

## 輸入資料規格

必要欄位（欄名完全一致）：
- `Customer Name` — 客戶名稱
- `Ship Date` — 出貨日期（容錯解析，NaT 自動跳過）
- `QTY` — 數量
- `SALES Total AMT` — 銷售額（TWD）
- `final GP(NTD,data from Financial Report)` — 毛利
- `Part Number` — 料號
- `Category` — 分類

選用欄位：
- `DES` — 用於 DES 關鍵字分類（若無此欄，DES 分類停用）

---

## Category 分類邏輯

優先順序：
1. Category 欄直接比對（Tablet / CDR / Tablet ACC / CDR ACC，大小寫不敏感）
2. DES 欄關鍵字比對（DES_RULES 字典，substring contains）
3. Fallback → Others

有效 Category：Tablet / CDR / Tablet ACC / CDR ACC / AI_SW / Others

DES_RULES（修改時需同步更新 app.py 頂部字典）：
- CDR ACC: cdr, gemini, evo, sprint, sd card, panic button, iosix, uvc camera,
           k220, k245, k265, smart link dongle, safetycam
- Tablet ACC: tablet, prometheus, chiron, hera, phaeton, surfing pro, cradle,
              f840, ulmo, fleet cable
- AI_SW: visionmax

---

## 關鍵設計決策

- **overrides.json**: Key 為 (Customer Name, Part Number, Month, DES) 的複合
  key，避免 Excel 更新後 index 偏移。跨 session / 重啟保留。
- **Cache busting**: DES_RULES 變更時透過 _rules_key() 自動使
  @st.cache_data 失效。
- **--server.headless true**: launcher.py 控制開瀏覽器時機（偵測 port 就緒
  再開），不依賴 Streamlit 預設行為。
- **更新 app.py 不需重新打包**: 直接替換 dist/啟動程式/app/app.py 即可。

---

## 目前版本

v3.1（最新）— 完整功能建置完成。
當前開發第一優先項：Windows Launcher 打包實作（PyInstaller）。

---

## 常見工作模式

- 修改分類規則 → 編輯 app.py 內的 DES_RULES，並同步更新 Notion 對照表
- 新功能開發 → py -m streamlit run app/app.py
- 出貨給使用者 → 執行 build.bat 重新打包
- 小修正（只改 app.py）→ 直接替換檔案，不需重新打包
