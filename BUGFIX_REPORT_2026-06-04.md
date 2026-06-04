# SMS 学生成绩上传工具 - Bug 修复报告
**日期**: 2026年6月4日  
**版本**: v2.1（修复版）

---

## 问题描述

在您的账号下，project_input_page.py 第537行的"**项目列表**"会根据第492行的"**活动代码**"选择动态改变。但同事的账号无法动态更新项目列表，完全没有反应。

---

## 根本原因分析

当您选择"活动代码"时，`project_input_page.py` 第532行的信号连接会触发 `_on_selection_changed()` 方法，该方法需要执行以下步骤：

1. 获取已保存的凭证（账号和密码）
2. 使用凭证连接到 SMS 系统
3. 根据选择的单位和活动代码搜索旧项目
4. 更新项目列表

**问题所在**：当同事第一次运行您的应用时，他们的账号下没有保存凭证。原代码在凭证为空时**直接返回，没有任何提示**，导致同事看不到任何反应。

### 代码位置
- **文件**: `sms_app/ui/pages/project_input_page.py`
- **方法**: `_on_selection_changed()`
- **第652-653行**: 原代码问题所在

---

## 修复方案

### 修复内容

修改 `_on_selection_changed()` 方法，当凭证为空时显示明确的警告消息：

**修改前**（第652-653行）：
```python
if not username or not password:
    return
```

**修改后**：
```python
if not username or not password:
    self.console.log_warning("❌ 未找到保存的凭证，请先在【设定】页面保存账号密码")
    self.old_projects_table.setRowCount(0)
    return
```

### 修复效果

现在当同事选择"活动代码"时，如果没有保存凭证，会在主窗口的**日志控制台**看到以下信息：

```
❌ 未找到保存的凭证，请先在【设定】页面保存账号密码
```

这样就能清楚地告诉用户需要做什么。

---

## 使用新版本的步骤

### 对于同事

1. **关闭旧版本**的应用
2. **删除旧版本**的 `SMS成绩上传工具.exe`
3. **下载新版本** `SMS成绩上传工具.exe`（已重新打包）
4. **首次启动**时，进行以下操作：
   - 点击左侧导航栏中的**【设定】**选项卡
   - 在**账号**和**密码**字段中输入 SMS 系统的凭证
   - 点击**保存设置**按钮
5. **返回【项目设置】**选项卡
6. 选择**单位**和**活动代码**，项目列表应该自动更新

### 对于您

您可以直接使用新版本，不需要重新配置凭证（已保存的凭证仍然有效）。

---

## 修复后的代码位置

- **文件**: `sms_app/ui/pages/project_input_page.py`
- **方法**: `_on_selection_changed()`  
- **第648-658行**: 已修复的代码

```python
def _on_selection_changed(self):
    """单位或活动代码改变时触发 - 自动搜索并更新旧项目列表"""
    prefix_code = self._get_prefix_code()
    if not prefix_code:
        return
    
    # 获取保存的凭证
    username, password = self.config.get_credentials()
    if not username or not password:
        self.console.log_warning("❌ 未找到保存的凭证，请先在【设定】页面保存账号密码")
        self.old_projects_table.setRowCount(0)
        return
    
    self.console.log_info(f"正在搜索项目: {prefix_code}...", "#dcdcaa")
    
    # 启动后台线程
    self.search_thread = SearchProjectThread(username, password, prefix_code)
    self.search_thread.search_finished.connect(self._on_search_finished)
    self.search_thread.start()
```

---

## 新版本信息

- **版本号**: v2.1（修复版）
- **打包日期**: 2026年6月4日 上午 10:34:27
- **打包文件**: `dist/SMS成绩上传工具.exe`

---

## 附加说明

### 凭证安全性
- 所有凭证都被加密存储在用户的本地目录 `~/.sms_app/` 中
- 每个用户账号都有独立的加密密钥
- 凭证**不会**包含在应用的打包文件中

### 如果再次出现类似问题
- 确保用户已在【设定】页面保存凭证
- 查看日志控制台中是否出现了警告消息
- 如果连接失败，检查账号和密码是否正确

---

## 总结

这个修复确保了当用户还没有保存凭证时，会立即收到清晰的反馈，而不是看到应用"完全没有动静"。这大大改善了用户体验，并使问题排查变得更加容易。

✅ **修复已完成并重新打包**
