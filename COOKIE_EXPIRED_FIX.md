# Cookie 过期自动恢复功能说明

## 问题描述
在执行成绩上传或项目添加时，如果遇到登入失败问题（通常是因为 Cookie 已过期），系统现在会自动重新获取新的 Session 并继续执行操作。

## 实现方案

### 1. 核心改进 - SMSHandler 中的 `upload_student_scores` 方法

**文件**: `sms_app/core/sms_handler.py`

新增方法 `upload_student_scores()` 具有以下特性：

```python
def upload_student_scores(self, username: str, password: str, scores_data: list, 
                         max_retries: int = 3, retry_delay: int = 2) -> dict:
```

**功能**：
- ✅ 自动登入系统
- ✅ 检测登入失败（Cookie 过期）
- ✅ **自动重新登入**（支持多次重试，默认 3 次）
- ✅ 上传学生成绩数据
- ✅ 返回详细的上传结果

**返回值**：
```python
{
    'success': bool,          # 是否全部成功
    'uploaded': int,          # 成功上传的条数
    'failed': int,            # 失败的条数
    'total': int,             # 总条数
    'message': str,           # 结果消息
    'errors': list            # 错误详情列表
}
```

### 2. AddProjectThread 改进

**文件**: `sms_app/ui/pages/project_input_page.py`

`AddProjectThread` 现在：
- ✅ 支持最多 3 次重试
- ✅ 在登入失败时自动重试
- ✅ 在添加项目失败时自动重试（可能是 Cookie 过期）
- ✅ 每次重试间隔 2 秒

```python
class AddProjectThread(QThread):
    """后台线程添加项目（支持自动重新登入）"""
```

### 3. UploadThread 完整实现

**文件**: `sms_app/ui/pages/score_upload_page.py`

`UploadThread.run()` 现在：
- ✅ 读取 Excel 文件中的学生数据
- ✅ 调用 `upload_student_scores()` 方法
- ✅ 自动处理登入失败和 Cookie 过期
- ✅ 提供进度反馈
- ✅ 返回详细的上传结果

```python
class UploadThread(QThread):
    """后台线程执行上传（支持自动重新登入）"""
    
    def run(self):
        # 1. 读取 Excel 文件
        # 2. 调用 upload_student_scores() 上传
        # 3. 如果 Cookie 过期，自动重新登入并重试
        # 4. 返回上传结果
```

## 工作流程

### 执行流程图

```
开始上传
  ↓
读取 Excel 数据
  ↓
尝试第 1 次上传
  ├─ 登入成功 → 上传数据 → 成功 ✅ → 返回结果
  ├─ 登入成功 → 上传数据 → 失败（Cookie过期）
  │   ↓
  │ 检测到 Cookie 过期错误
  │   ↓
  │ 关闭驱动，重新初始化
  │   ↓
  │ 等待 2 秒
  │   ↓
  │ 尝试第 2 次上传 (重复上述流程)
  │
  └─ 登入失败
      ↓
    等待 2 秒
      ↓
    尝试第 2 次登入
      ↓
    ... (最多重试 3 次)
      ↓
    如果仍然失败，返回错误结果
```

## 使用方法

### 1. 成绩上传

在成绩上传页面：

```
1. 选择 Excel 文件 (包含学生数据)
2. 点击 "📤 开始上传"
3. 系统自动：
   - 读取数据
   - 尝试登入
   - 如果 Cookie 过期，自动重新登入（最多 3 次）
   - 上传数据并显示进度
```

### 2. 项目添加

在项目输入页面：

```
1. 填写项目信息（代码、名称等）
2. 点击 "➕ 添加到 SMS"
3. 系统自动：
   - 尝试登入
   - 如果 Cookie 过期，自动重新登入（最多 3 次）
   - 添加项目
```

## 技术细节

### Cookie 过期检测

系统通过以下方式检测 Cookie 过期：

1. **登入URL检查**：如果页面仍在登入页面，说明登入失败
2. **异常信息检查**：检查是否包含 `'login'`, `'session'`, `'unauthorized'` 等关键词
3. **页面加载失败**：如果无法找到预期的表单元素

### 自动恢复机制

当检测到 Cookie 过期时：

1. 关闭当前浏览器驱动
2. 重新初始化驱动程序
3. 等待 2 秒（给服务器反应时间）
4. 使用相同凭证重新登入
5. 继续之前的操作

### 重试参数

可以在调用时自定义：

```python
# 最多重试 5 次，每次间隔 3 秒
result = handler.upload_student_scores(
    username=username,
    password=password,
    scores_data=scores_data,
    max_retries=5,    # 修改重试次数
    retry_delay=3     # 修改重试延迟（秒）
)
```

## 日志输出

系统会输出详细的日志信息：

```
📍 第 1 次尝试登入...
📍 打开登入页面: http://sms.chhsban.edu.my/sms/index.php?r=site/login
✓ 页面已加载
⏳ 等待登入表单加载...
✓ 登入表单已加载
📝 输入帐号: your_username
✓ 帐号已输入
🔐 输入密码
✓ 密码已输入
🖱️  点击登入按钮...
✓ 按钮已点击
⏳ 等待登入完成...
✓ 登入成功！当前 URL: ...
✓ 登入成功，开始上传 5 条成绩数据...
  📤 上传 [1/5] 12A 001 张三 ✅
  📤 上传 [2/5] 12A 002 李四 ✅
  ...
✅ 上传完成: 5 成功, 0 失败
```

## 错误处理

如果所有重试都失败，系统会：

1. 输出详细的错误信息
2. 在结果中标记为 `success: False`
3. 提供错误列表供用户查看
4. 显示用户友好的错误消息

## 测试建议

### 模拟 Cookie 过期

为了测试自动重新登入功能，可以：

1. **关闭浏览器**：在上传过程中关闭浏览器窗口，模拟 Session 丢失
2. **服务器重启**：如果有权限，重启 SMS 服务器来清空 Session
3. **长时间等待**：让程序等待足够长的时间使 Cookie 自然过期

### 验证日志

检查控制台输出，确认：

- ✅ 检测到登入失败
- ✅ 触发自动重新登入
- ✅ 重新登入成功
- ✅ 上传继续进行
- ✅ 最终上传成功

## 常见问题

### Q: 为什么要重试 3 次？
A: 3 次重试是一个平衡值。足够给予系统恢复的机会，但又不会太长时间等待。

### Q: 可以关闭自动重试吗？
A: 可以，在调用时设置 `max_retries=1`：
```python
result = handler.upload_student_scores(..., max_retries=1)
```

### Q: 如果密码已过期会怎样？
A: 第一次登入就会失败，重试也会失败。系统会在第一次失败后提示用户检查凭证。

### Q: 上传过程中断后能恢复吗？
A: 目前不支持断点续传。需要重新选择文件并重新上传。

## 后续改进建议

1. 📝 **断点续传**：记录已上传的行，支持从中断处继续
2. 🔄 **更智能的重试**：根据错误类型采用不同的重试策略
3. 📊 **上传日志**：保存每次上传的详细日志供查阅
4. 🔐 **Token刷新**：定期主动刷新 Token/Session，而不是等待过期

## 相关文件修改

- `sms_app/core/sms_handler.py`：新增 `upload_student_scores()` 方法
- `sms_app/ui/pages/score_upload_page.py`：完整实现 `UploadThread.run()`
- `sms_app/ui/pages/project_input_page.py`：改进 `AddProjectThread` 重试机制

## 联系和反馈

如遇到问题或有改进建议，请提供：
- 完整的错误日志输出
- 操作步骤
- 预期结果 vs 实际结果
