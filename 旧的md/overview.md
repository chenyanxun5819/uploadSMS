# 编程目的：上传学生成绩至SMS的特定页面

## 必须完成的要件及最终结果：
- 设定页面：可设定使用者帐号和密码，预设执行时为空值。
- 設定項目頁面。
- 批量上傳成績頁面。
- 最终必须打包成可执档。

## 設計風格：
- 仿vscode的介面設計。

### 啟動頁面。
總尺寸：1200*800
三個操作按鈕順序：項目輸入、成績輸入、設定
區域	            尺寸	        功能
上部左側輸入區	    1050 × 200	    成績輸入表單
上部左側預覽區	    1050 × 300	    成績預覽表格
上部右側按鈕區	    150 × 500	    三個操作按鈕
下部 Console 區	    1200 × 300	   即時執行結果／Log

### 設定：
- 記錄頁面登入需要的帳號、密碼，
- 登入頁面http://sms.chhsban.edu.my/sms/index.php?r=site/login
- username的完整Xpath：/html/body/div[2]/div/div/div/div[2]/div/form/div[2]/input
- password的完整Xpath：/html/body/div[2]/div/div/div/div[2]/div/form/div[3]/input
- 登入按鈕完整Xpath：/html/body/div[2]/div/div/div/div[2]/div/form/div[4]/div/button
- 此部份為編程的背景程式，「在程式啟動時，先嘗試使用瀏覽器保存的 Cookie 來取得伺服器 Session。如果 Session 還有效，就可以直接操作；如果 Session 已失效（例如過期或被清除），則必須重新執行登入流程。」并在『log區』中顯示執行動作。

### 項目輸入
- 記錄所需輸入的比賽項目。
- 輸入頁面連結：http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index
- 欄位一：
    - 名稱：项目代码
    - 對應欄位完整Xpath：/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[2]/div/input
- 欄位二：
    - 名稱：项目名称 
    - 對應欄位完整Xpath：/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[3]/div/input
- 欄位三：
    - 名稱：分数项目
    - 對應欄位完整Xpath：/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[4]/div/input
    - 預設值為0。
- 保存按鈕：/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[5]/button[1]


### 成績輸入
- 輸入"比賽項目"中的學生成績。
- 上部左側輸入區：
    - 下载excel样本：C:\Users\MSI\Documents\2025_affairs\學術結果登記\学术上传python\template.xlsx
    - 选取要输入的excel档。
        - 选取后，在"上部左側預覽區"中显示excel内容。
        - 并根据A2的活动代码，在下方显示项目名称。
    - 确认上传。
- 此部份请参考C:\Users\MSI\Documents\2025_affairs\學術結果登記\学术上传python\upload.py中的编码，大至都是相同的，唯将活动代码，由A2，改为B2，日期由A1改为B1。
