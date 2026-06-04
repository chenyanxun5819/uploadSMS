#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS 应用 - Headless 模式使用指南
如何启用/禁用后台执行模式
"""

print("""
╔══════════════════════════════════════════════════════════════════════════╗
║                                                                          ║
║              SMS 应用 Headless 模式使用指南                              ║
║               Enable/Disable Background Execution Mode                   ║
║                                                                          ║
╚══════════════════════════════════════════════════════════════════════════╝


📋 目录
═══════════════════════════════════════════════════════════════════════════
  1. 快速设置
  2. 配置位置
  3. 三种使用方法
  4. 示例代码
  5. 故障排查


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. 快速设置
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

现在系统支持以下模式：

✅ 显示浏览器（默认）
   • headless = False
   • 浏览器窗口可见
   • 可观察整个操作过程
   • 用于开发和调试

✅ 后台执行（新增）
   • headless = True
   • 浏览器窗口隐藏
   • 执行结果完全相同
   • 用于生产环境


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
2. 配置位置
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

配置文件位置：
  Windows:   C:\\Users\\<用户名>\\.sms_app\\config.json
  Linux:     /home/<用户名>/.sms_app/config.json
  macOS:     /Users/<用户名>/.sms_app/config.json

配置文件示例：
  {
    "credentials": {...},
    "session": {...},
    "projects": [...],
    "browser": {
      "headless": false      ← 改为 true 启用后台模式
    }
  }


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
3. 三种使用方法
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【方法 A】通过 Python 代码动态设置（推荐）
─────────────────────────────────────────────────

  from sms_app.core.config_manager import ConfigManager
  
  config = ConfigManager()
  
  # 启用后台模式
  config.set_headless_mode(True)
  print("✅ 已启用后台模式")
  
  # 禁用后台模式（显示浏览器）
  config.set_headless_mode(False)
  print("✅ 已禁用后台模式")
  
  # 检查当前模式
  is_headless = config.get_headless_mode()
  if is_headless:
      print("🔇 当前：后台模式")
  else:
      print("🔊 当前：显示浏览器")


【方法 B】在 UI 中切换（通过代码）
─────────────────────────────────────────────────

  # 在 settings_page.py 中添加切换开关
  # （可以在"设定"页面添加一个复选框）
  
  headless_checkbox = QCheckBox("后台模式（隐藏浏览器窗口）")
  headless_checkbox.setChecked(self.config.get_headless_mode())
  headless_checkbox.stateChanged.connect(
      lambda state: self.config.set_headless_mode(state == Qt.CheckState.Checked)
  )
  
  settings_layout.addWidget(headless_checkbox)


【方法 C】直接编辑配置文件
─────────────────────────────────────────────────

  1. 找到配置文件：
     C:\\Users\\<用户名>\\.sms_app\\config.json

  2. 编辑文件，将 headless 值改为：
     - true  → 后台执行（无窗口）
     - false → 显示浏览器

  3. 保存文件，重新启动应用


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
4. 示例代码
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【示例 1】在应用启动时设置后台模式
─────────────────────────────────────────────────

  # 在 sms_app/main_app.py 中

  from sms_app.core.config_manager import ConfigManager
  
  def main():
      app = QApplication(sys.argv)
      config = ConfigManager()
      
      # 启动时自动使用后台模式
      config.set_headless_mode(True)
      
      window = MainWindow()
      window.show()
      sys.exit(app.exec())


【示例 2】根据环境变量切换模式
─────────────────────────────────────────────────

  import os
  from sms_app.core.config_manager import ConfigManager
  
  config = ConfigManager()
  
  # 从环境变量读取，默认为 False
  headless = os.getenv('SMS_HEADLESS', 'false').lower() == 'true'
  config.set_headless_mode(headless)
  
  # 使用方式：
  # SMS_HEADLESS=true python sms_app/main_app.py    # 后台模式
  # SMS_HEADLESS=false python sms_app/main_app.py   # 显示浏览器


【示例 3】命令行工具切换模式
─────────────────────────────────────────────────

  # 创建一个命令行工具脚本：sms_config.py

  import argparse
  from sms_app.core.config_manager import ConfigManager

  def main():
      parser = argparse.ArgumentParser(description='SMS 应用配置工具')
      parser.add_argument('--headless', choices=['on', 'off'], 
                          help='启用或禁用后台模式')
      
      args = parser.parse_args()
      config = ConfigManager()
      
      if args.headless:
          headless = args.headless == 'on'
          config.set_headless_mode(headless)
          status = "✅ 已启用后台模式" if headless else "✅ 已禁用后台模式"
          print(status)
      else:
          current = "后台模式" if config.get_headless_mode() else "显示浏览器"
          print(f"当前模式: {current}")

  if __name__ == '__main__':
      main()

  # 使用方式：
  # python sms_config.py                # 查看当前模式
  # python sms_config.py --headless on  # 启用后台
  # python sms_config.py --headless off # 禁用后台


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
5. 故障排查
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Q: 后台模式下无法看到浏览器，如何调试？
A: 可以改回显示模式：
   config.set_headless_mode(False)
   然后观察浏览器窗口的操作过程

Q: 后台模式运行比较慢，怎么办？
A: 后台模式理论上应该更快。如果慢的话，可能是：
   • 网络连接不稳定
   • 系统资源不足
   • 检查网络是否正常

Q: 如何在脚本中同时使用两种模式？
A: 可以创建两个线程，分别设置不同的 headless 值

Q: 能否在运行时切换模式？
A: 当前设计中，需要重启应用才能切换
   但可以修改代码支持热切换（较复杂）


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
立即测试
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. 启用后台模式：
   python -c "
   from sms_app.core.config_manager import ConfigManager
   config = ConfigManager()
   config.set_headless_mode(True)
   print('✅ 后台模式已启用')
   "

2. 禁用后台模式（恢复显示浏览器）：
   python -c "
   from sms_app.core.config_manager import ConfigManager
   config = ConfigManager()
   config.set_headless_mode(False)
   print('✅ 后台模式已禁用')
   "

3. 查看当前设置：
   python -c "
   from sms_app.core.config_manager import ConfigManager
   config = ConfigManager()
   mode = '后台模式' if config.get_headless_mode() else '显示浏览器'
   print(f'当前: {mode}')
   "


═══════════════════════════════════════════════════════════════════════════
                              总结
═══════════════════════════════════════════════════════════════════════════

✅ 系统现已支持 Headless 模式
   • 可在后台执行，无需显示浏览器
   • 支持通过代码、配置文件等多种方式切换
   • 开发时使用显示模式，生产时使用后台模式

✅ 三种配置方法任选其一
   • 方法 A：Python 代码（最灵活）
   • 方法 B：UI 切换（最方便）
   • 方法 C：编辑配置文件（最直接）

✅ 使用 ChromeDriver 是必需的
   • SMS 系统是 JavaScript 动态网站
   • 需要真实浏览器环境

═══════════════════════════════════════════════════════════════════════════
""")
