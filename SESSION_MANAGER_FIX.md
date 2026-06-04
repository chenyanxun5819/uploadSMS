# SMS 應用修復記錄 - Session 重用優化

## 問題分析

在 sms_app_2026-06-04.log 第52行發現：
- 新增項目成功：✅ 項目已添加: ACA CMO209
- 但隨後獲取最後一條記錄失敗：❌ 登入失敗

### 根本原因

1. **啟動時**（第8行）：`StartupChecker` 建立第一個session進行驗證
2. **新增項目時**：`AddProjectThread` 建立**新的獨立session**進行登入
3. **獲取記錄時**：`FetchLastProjectThread` 又建立**第三個獨立session**進行登入
   - 此時第二次登入失敗！

### 為什麼登入失敗？

- 每個線程都創建了獨立的 `requests.Session()` 對象
- 沒有重用已有的 cookies / session
- 導致系統認為是不同的客戶端重複登入，可能觸發安全限制

### 正確的術語

- **Session**：`requests.Session()` 對象，包含了cookies、連接池等
- **Cookie**：HTTP cookies，session會自動存儲並重用
- 正確做法：**重用已有的 session 對象**（而不是每次創建新的）

## 解決方案

### 1. 創建全局 Session 管理器
- 新文件：`sms_app/core/session_manager.py`
- 使用**單例模式**（Singleton）確保全局只有一個session
- 提供認證狀態管理，避免重複登入

### 2. 修改 AddProjectThread
- 使用全局session而不是創建新的
- 檢查認證狀態，如已認證則跳過登入
- 成功後設置 `is_authenticated = True`

### 3. 修改 FetchLastProjectThread
- 使用全局session
- 檢查認證狀態，如已認證則直接使用
- 無需重新登入

### 4. 修改 StartupChecker
- 使用全局session而不是創建新的
- 檢查認證狀態，如已認證則重用
- 完成後不關閉全局session（保留用於後續操作）

## 修改的文件

| 文件 | 修改內容 | 重要性 |
|-----|--------|------|
| `sms_app/core/session_manager.py` | 新建全局session管理器 | 🔴 關鍵 |
| `sms_app/ui/pages/project_input_page.py` | AddProjectThread + FetchLastProjectThread 使用全局session | 🔴 關鍵 |
| `sms_app/core/startup_checker.py` | check_and_update使用全局session | 🟠 重要 |

## 工作流程改進

### 原來的流程
```
啟動時:    session1 → 登入 → 驗證 → 關閉
新增項目:  session2 → 登入 → 添加 → 完成
獲取記錄:  session3 → 登入❌失敗
```

### 修復後的流程
```
啟動時:    全局session → 登入 → 驗證 → 保留
新增項目:  全局session → 已認證，跳過登入 → 添加 → 完成
獲取記錄:  全局session → 已認證，跳過登入 → 獲取 → 成功
```

## 預期效果

✅ 不再出現重複登入錯誤  
✅ 新增項目後能正常獲取記錄  
✅ 項目列表能正常更新  
✅ 應用啟動速度提升（減少登入次數）  
✅ 更穩定的session管理  

## 驗證方法

1. 重新打包應用
2. 啟動應用，觀察日誌：
   - 應該看到：「✅ 使用現有會話，無需重新登入」
   - 不應該再看到：「❌ 登入失敗」
3. 測試新增項目流程
4. 驗證項目列表是否更新
