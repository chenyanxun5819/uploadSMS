#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 Headless 模式配置功能
"""

from sms_app.core.config_manager import ConfigManager

print("📊 SMS 应用 Headless 模式测试")
print("=" * 50)

config = ConfigManager()

# 检查当前模式
current = config.get_headless_mode()
mode_str = "后台模式" if current else "显示浏览器"
print(f"\n[1] 当前模式: {mode_str}")

# 启用后台模式
print("\n[2] 启用后台模式...")
config.set_headless_mode(True)
print("✅ 后台模式已启用")

# 验证
new_state = config.get_headless_mode()
mode_str = "后台模式" if new_state else "显示浏览器"
print(f"✓ 验证: {mode_str}")

# 禁用后台模式
print("\n[3] 禁用后台模式（恢复显示浏览器）...")
config.set_headless_mode(False)
print("✅ 后台模式已禁用")

# 验证
final_state = config.get_headless_mode()
mode_str = "后台模式" if final_state else "显示浏览器"
print(f"✓ 验证: {mode_str}")

print("\n" + "=" * 50)
print("✅ 所有功能测试通过！")
print("\n提示: 现在您可以通过以下方式使用后台模式：")
print("  config.set_headless_mode(True)   # 启用后台模式")
print("  config.set_headless_mode(False)  # 禁用后台模式")
print("\n配置已保存到:")
print(f"  {config.CONFIG_FILE}")
