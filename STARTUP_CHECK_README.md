# SMS 学生成绩系统 - 启动检查功能说明

## 功能概述

在应用启动时自动检查项目数据是否与服务器同步，如果不同步则自动更新。

## 功能特点

✅ **自动检查** - 每次启动时自动检查  
✅ **智能更新** - 支持增量更新，只下载新增项目  
✅ **实时日志** - 在控制台实时显示检查和更新过程  
✅ **无感更新** - UI不阻塞，后台进行检查  
✅ **完整记录** - 所有操作都被记录到日志文件

## 工作流程

```
应用启动
    ↓
显示主窗口
    ↓
延迟 500ms（等待 UI 准备）
    ↓
执行启动检查
    ├─ 检查凭证是否已保存
    ├─ 连接到 SMS 系统
    ├─ 获取页面总数
    ├─ 与缓存比对
    │  ├─ 一致 ✅ → 无需更新
    │  └─ 不一致 ⚠️ → 进行更新
    │      ├─ 首次导入 → 下载全部项目
    │      └─ 增量更新 → 只下载新项目
    └─ 显示结果到 Console
```

## 日志输出说明

### 成功消息（绿色 ✅）
```
[✓] SMS 学生成绩自动上传系统已启动
[✓] 已连接
[✓] 页面总数: 2530
[✓] 数据一致，无需更新
[✓] 缓存更新成功
[✓] 数据检查完成 - ✅ 数据一致 (总数: 2530)
```

### 警告消息（黄色 ⚠️）
```
[⚠] 未保存凭证，跳过检查
[⚠] 数据不一致！
[⚠] 数据检查跳过 - 未保存凭证
```

### 错误消息（红色 ❌）
```
[✗] 连接失败: ...
[✗] 无法从页面获取总数
[✗] 获取新增项目失败: ...
[✗] 缓存更新失败: ...
```

### 信息消息（蓝色 ℹ️）
```
[INFO] 📍 正在连接系统...
[INFO] 📍 获取页面总数...
[INFO] 📍 首次导入，下载所有 2530 条项目...
[INFO] ℹ️ 合并后共 2530 条项目
```

## 关键 AJAX 端点

项目数据通过以下 AJAX 端点获取和更新：
```
http://sms.chhsban.edu.my/sms/index.php?ItemM_page={page}&ajax=item-m-grid&r=transaction/itemSetting/index
```

参数说明：
- `ItemM_page`: 分页页码（从 1 开始）
- `ajax=item-m-grid`: 指示返回 HTML 表格片段
- `r=transaction/itemSetting/index`: 路由参数

系统会从 AJAX 响应中提取：
- 分页信息："第 X-Y 条，共 Z 条"
- 项目表格数据：序号、代码、名称、分数

## 检查结果代码

启动检查返回的结果包含以下字段：

```python
{
    'checked': bool,      # 是否检查成功
    'page_total': int,    # 页面上的总数
    'cached_total': int,  # 缓存中的总数
    'matched': bool,      # 是否匹配
    'updated': bool,      # 是否已更新
    'message': str        # 详细消息
}
```

### 检查结果示例

**数据一致（无需更新）**
```python
{
    'checked': True,
    'page_total': 2530,
    'cached_total': 2530,
    'matched': True,
    'updated': False,
    'message': '✅ 数据一致 (总数: 2530)'
}
```

**数据不一致（已更新）**
```python
{
    'checked': True,
    'page_total': 2535,
    'cached_total': 2530,
    'matched': False,
    'updated': True,
    'message': '✅ 已更新 (2530 → 2535，新增 5 条)'
}
```

**凭证缺失（跳过检查）**
```python
{
    'checked': False,
    'page_total': 0,
    'cached_total': 0,
    'matched': False,
    'updated': False,
    'message': '未保存凭证'
}
```

## 缓存位置

项目缓存文件存储在：
- **缓存目录**: `~/.sms_app/`
- **项目文件**: `~/.sms_app/projects.json`
- **元数据文件**: `~/.sms_app/metadata.json`
- **日志文件**: `~/.sms_app/logs/sms_app_YYYY-MM-DD.log`

在 Windows 上，`~` 通常是 `C:\Users\<username>`

## 手动测试

可以使用提供的测试脚本进行手动测试：

```bash
cd 学术上传python
python test_startup_checker.py
```

## 相关文件

- **核心检查器**: `sms_app/core/startup_checker.py`
- **主应用入口**: `sms_app/main_app.py`
- **缓存管理器**: `sms_app/core/cache_manager.py`
- **凭证管理器**: `sms_app/core/config_manager.py`
- **测试脚本**: `test_startup_checker.py`

## 配置说明

启动检查的延迟时间可以在 `main_app.py` 中调整：

```python
timer.start(500)  # 延迟 500ms 执行，可改为其他值（毫秒）
```

## 常见问题

### Q: 为什么没有显示检查日志？
A: 检查需要已保存的凭证。请先在"设置"页面输入用户名和密码。

### Q: 如何手动触发检查？
A: 目前检查只在应用启动时执行。如需手动检查，可运行测试脚本。

### Q: 更新需要多长时间？
A: 取决于项目总数和网络速度。首次导入较慢，增量更新较快。

### Q: 是否会阻塞主界面？
A: 不会。检查在后台进行，不影响 UI 响应。

## 更新日志

### v1.0 (2024)
- ✅ 实现启动自动检查
- ✅ 支持增量更新
- ✅ 实时日志输出
- ✅ 完整错误处理
