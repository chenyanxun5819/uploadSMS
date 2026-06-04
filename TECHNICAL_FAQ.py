#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS 系统技术FAQ - 关于 ChromeDriver 和后台执行
Q&A: ChromeDriver Usage and Background Execution
"""

print("""
╔══════════════════════════════════════════════════════════════════════════╗
║                        SMS 系统技术 FAQ                                 ║
║               ChromeDriver 和后台执行的技术说明                          ║
╚══════════════════════════════════════════════════════════════════════════╝

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
问题 1：这里还是使用 ChromeDriver 吗？
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ 是的，仍然使用 ChromeDriver

【当前架构】
  
  project_input_page.py (UI 层)
         ↓
  AddProjectThread (后台线程)
         ↓
  SMSHandler (Selenium 封装层)
         ↓
  WebDriver + ChromeDriver (浏览器驱动)
         ↓
  Google Chrome 浏览器
         ↓
  SMS 系统网站


【为什么必须使用 ChromeDriver？】

  SMS 系统网站是 JavaScript 渲染的动态网站：
  
  ✗ 使用 requests 库无法工作
    原因：requests 只能获取静态 HTML，无法执行 JavaScript
  
  ✗ 需要真实浏览器环境
    原因：填写表单、点击按钮等操作需要真实的浏览器引擎
  
  ✅ ChromeDriver 提供真实浏览器支持
    优势：完整的 JavaScript 支持、表单操作、Cookie 管理


【技术依赖】

  安装的包：
  ├─ selenium (WebDriver 框架)
  ├─ webdriver-manager (自动下载 ChromeDriver)
  ├─ PyQt6 (UI 界面)
  └─ requests (仅用于备用方案，主要使用 Selenium)


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
问题 2：一定要开启 Web？能否在后台执行？
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ 可以在后台执行！使用 Headless 模式


【当前状态】

  现在：headless=False
  ├─ 浏览器窗口可见
  ├─ 可以观察执行过程（用于调试）
  ├─ 用户体验：显示操作进度
  └─ 性能：稍微慢一些（GPU 绘制开销）

  可改为：headless=True
  ├─ 浏览器在后台运行（无窗口）
  ├─ 执行结果相同
  ├─ 用户体验：无界面干扰
  └─ 性能：更快（无 GPU 开销）


【切换到 Headless 模式的方法】

  方案 A：直接修改代码（推荐）
  ─────────────────────────────────────────────────
  
  文件：sms_app/ui/pages/project_input_page.py
  
  找到这一行：
    handler = SMSHandler(headless=False)
  
  改为：
    handler = SMSHandler(headless=True)


  方案 B：添加配置选项（更灵活）
  ─────────────────────────────────────────────────
  
  在 ConfigManager 中添加：
    class ConfigManager:
        def __init__(self):
            ...
            self.headless_mode = True  # 新增配置
        
        def set_headless(self, value: bool):
            self.headless_mode = value
  
  在 project_input_page.py 中使用：
    headless = self.config.get_headless_mode()
    handler = SMSHandler(headless=headless)


  方案 C：环境变量控制（最灵活）
  ─────────────────────────────────────────────────
  
  import os
  headless = os.getenv('SMS_HEADLESS', 'True').lower() == 'true'
  handler = SMSHandler(headless=headless)
  
  使用：
    SMS_HEADLESS=False python sms_app/main_app.py  # 显示浏览器
    SMS_HEADLESS=True python sms_app/main_app.py   # 后台运行


【两种模式的优缺点比较】

  模式             | headless=False    | headless=True
  ────────────────────────────────────────────────────────
  浏览器窗口       | ✅ 可见            | ✗ 隐藏
  执行速度        | ⚠️ 中等            | ✅ 较快
  调试方便性      | ✅ 容易观察        | ⚠️ 难观察
  资源占用        | ⚠️ 较多            | ✅ 较少
  在后台运行      | ⚠️ 受阻碍          | ✅ 完全支持
  稳定性          | ✅ 高              | ⚠️ 偶尔有兼容性
  用户体验        | ⚠️ 窗口闪烁        | ✅ 静默
  CI/CD 部署      | ✗ 不适合           | ✅ 最佳


【推荐方案】

  ╔─────────────────────────────────────────────────╗
  │  场景 1：本地开发和调试                         │
  │    → 使用 headless=False                        │
  │    → 可以看到浏览器在做什么                     │
  │    → 方便排查问题                               │
  │                                                 │
  │  场景 2：生产环境/后台服务                      │
  │    → 使用 headless=True                         │
  │    → 完全后台运行，无界面                       │
  │    → 性能更好                                   │
  │                                                 │
  │  场景 3：自动化部署/定时任务                    │
  │    → 使用 headless=True                         │
  │    → 适合 cron 任务或容器化部署                 │
  ╚─────────────────────────────────────────────────╝


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
实施建议
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【推荐实施方案：添加可配置的 Headless 模式】

  修改位置：
  
  1. sms_app/core/sms_handler.py
     ✅ 已支持 headless 参数（无需改动）
  
  2. sms_app/ui/pages/project_input_page.py
     需要修改：
       # 原：handler = SMSHandler(headless=False)
       # 改：headless = getattr(self.config, 'headless', False)
       #     handler = SMSHandler(headless=headless)
  
  3. sms_app/core/config_manager.py
     需要添加：
       def get_headless_mode(self) -> bool:
           return self.config.get('headless', False)
       
       def set_headless_mode(self, value: bool):
           self.config['headless'] = value


【立即实现的代码变更】

  执行以下修改，让用户可以选择是否显示浏览器窗口：
  
  ① 在 AddProjectThread 中添加参数
  ② 在 project_input_page.py 中读取配置
  ③ 在 config_manager.py 中保存选项


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FAQ 补充
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Q: Headless 模式下上传失败怎么办？
A: 大多数情况下 Headless 模式工作正常。如果失败，可以：
   1. 在测试阶段使用 headless=False 观察问题
   2. 检查网络连接是否正常
   3. 查看错误日志寻找具体原因

Q: 能否使用其他驱动程序（如 Firefox, Edge）？
A: 可以，Selenium 支持多个驱动程序。但需要：
   1. 安装对应的驱动程序
   2. 修改 SMSHandler 中的 webdriver 初始化代码
   3. 测试兼容性

Q: Headless 模式下还能看到进度吗？
A: 不能看到浏览器窗口，但可以通过：
   1. 控制台日志（print 输出）
   2. 日志文件
   3. UI 中的进度提示
   来了解执行进度

Q: 后台运行时能否被中断？
A: 可以。通过：
   1. Ctrl+C 中断
   2. 杀死进程
   3. 通过 UI 取消按钮（如果支持）


═══════════════════════════════════════════════════════════════════════════
                              总结
═══════════════════════════════════════════════════════════════════════════

✅ 当前系统使用 ChromeDriver
   → 这是必需的，因为 SMS 系统是 JavaScript 网站

✅ 可以在后台执行
   → 通过设置 headless=True 实现
   → 推荐在生产环境中使用

✅ 两种模式都完全支持
   → headless=False：适合开发调试
   → headless=True：适合生产环境

═══════════════════════════════════════════════════════════════════════════
""")
