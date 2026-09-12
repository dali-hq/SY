### 檔案說明
- `index.html` - 首頁
- `dayoff.html` - 請假登錄表單
- `duty.html` - 每月應休天數表
- `data.js` - 員工基本資料（含照片路徑）
- `employees.js` - 員工所屬工地（依前月臉部辨識）

### 資料分離設計
- **固定資料**（data.js）：工號、姓名、照片
- **變動資料**（employees.js）：所屬工地（調動式）
