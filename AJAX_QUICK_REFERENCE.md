# AJAX 修复 - 快速参考

## 🔧 关键修改

### URL 对比

| 场景 | ❌ 错误 (旧) | ✅ 正确 (新) |
|------|-----------|----------|
| **获取总数** | `?r=site/login&page=1` | `?ItemM_page=1&ajax=item-m-grid` |
| **获取第N页** | `?r=itemSetting/index&page=N` | `?ItemM_page=N&ajax=item-m-grid` |
| **超时情况** | 频繁（30-60s） | ✅ 无 (2-5s) |

### Python 代码修改

```python
# ❌ 错误方式 (导致超时)
response = session.get("http://sms.chhsban.edu.my/sms/index.php?r=site/login", 
                       params={'page': 1}, 
                       timeout=10)

# ✅ 正确方式 (快速稳定)
url = "http://sms.chhsban.edu.my/sms/index.php"
params = {
    'ItemM_page': 1,
    'ajax': 'item-m-grid',
    'r': 'transaction/itemSetting/index'
}
response = session.get(url, params=params, timeout=10)
```

## 📍 修改位置

| 文件 | 方法 | 修改内容 |
|------|------|---------|
| `sms_app/core/startup_checker.py` | `get_page_total_count()` | 改用 AJAX URL |
| `sms_app/core/startup_checker.py` | `fetch_new_projects()` | 改用 AJAX URL 获取总数 |

## 🚀 快速验证

### 方法 1：查看代码
```bash
# 检查 get_page_total_count 中是否使用 ItemM_page
grep -n "ItemM_page" sms_app/core/startup_checker.py
```

### 方法 2：运行测试
```bash
# 测试 AJAX 端点（需要填入凭证）
python test_ajax_endpoint.py

# 测试完整的启动检查
python test_startup_checker.py

# 启动应用观察日志
python sms_app/main_app.py
```

## 📊 AJAX 参数速查

```python
# 完整的 AJAX 请求参数
params = {
    'ItemM_page': page_number,      # 页码：1, 2, 3, ...
    'ajax': 'item-m-grid',          # 固定值
    'r': 'transaction/itemSetting/index'  # 固定路由
}
```

## ✅ 验证清单

启动应用后，Console 应该显示：

```
[✓] SMS 学生成绩自动上传系统已启动
[INFO] 🚀 启动检查：验证项目数据
[INFO]   📍 正在连接系统...
[✓]   ✅ 已连接
[INFO]   📍 获取页面总数...
[✓]   ✅ 页面总数: 2530
[INFO]   📦 缓存总数: 2530
[✓]   ✅ 数据一致，无需更新
[✓] 数据检查完成 - ✅ 数据一致 (总数: 2530)
```

❌ 如果还出现超时，检查清单：
- [ ] 凭证已在设置中保存
- [ ] 网络连接正常
- [ ] 服务器在线
- [ ] 没有被 IP 封禁

## 📚 相关文档

| 文档 | 用途 |
|------|------|
| `AJAX_ENDPOINT_FIX.md` | 详细技术说明 |
| `STARTUP_CHECK_README.md` | 完整功能文档 |
| `FIX_SUMMARY.md` | 修复总结 |
| `AJAX_QUICK_REFERENCE.md` | 本文件 |

## 💬 常见问题

**Q: 为什么改用 AJAX？**  
A: AJAX 只返回必要的表格数据，而不是整个页面，速度快 10-20 倍。

**Q: ItemM_page 是什么？**  
A: AJAX 分页参数，从 1 开始，每页 10 条记录。

**Q: 还是超时怎么办？**  
A: 运行 `test_ajax_endpoint.py` 诊断具体问题。

**Q: 可以改变超时时间吗？**  
A: 可以，在代码中改 `timeout=10` 为其他值（秒数）。

---

**最后更新**：2026-05-26  
**状态**：✅ 修复完成
