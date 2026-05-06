STARLUX EMD SOP 格式檢查系統
=============================
依照 SOP-EMD-22-101 Rev.05 自動驗證文件格式

【安裝需求】
Python 3.8+

安裝依賴套件：
  pip install flask python-docx

【啟動方式】
在 sop-checker 資料夾內執行：
  python app.py

瀏覽器開啟：
  http://localhost:7860

【檔案說明】
  app.py      Flask 後端伺服器
  checker.py  格式檢查引擎（規則邏輯皆在此修改）
  templates/  前端介面 HTML

【未來規範更新】
如需修改檢查規則，請編輯 checker.py：
  - DEPT_CODES：部門代碼清單
  - TIER2_NAMES / TIER3_NAMES：文件稱謂
  - REQUIRED_SECTIONS：必要章節
  - COVER_FIELDS：封面必要欄位

【可自動修改項目】
  ✅ 英文字體（Arial 12）
  ✅ 中文字體（微軟正黑體 12）
  ✅ 日期格式（→ MMM/DD/YYYY）

【需人工修改項目】
  ⚠️ 頁首/頁尾欄位內容
  ⚠️ 頁邊距
  ⚠️ 中英文排列順序
  ⚠️ 修訂紀要版本數
  ⚠️ 英文大小寫
  ⚠️ 文件稱謂/編號不符
