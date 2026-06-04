#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
启动检查功能测试脚本
"""

import sys
from pathlib import Path

# 添加项目路径
project_root = Path(__file__).parent / "sms_app"
sys.path.insert(0, str(project_root))

from core.startup_checker import StartupChecker


def test_startup_check():
    """测试启动检查功能"""
    
    print("\n" + "="*80)
    print("🧪 SMS 学生成绩系统 - 启动检查功能测试")
    print("="*80 + "\n")
    
    # 创建检查器
    checker = StartupChecker()
    
    # 收集日志
    logs = []
    
    def log_callback(message):
        """日志回调函数"""
        logs.append(message)
        print(message)
    
    # 执行检查
    print("\n📍 执行检查...\n")
    result = checker.check_and_update(log_callback=log_callback)
    
    # 显示结果
    print("\n" + "="*80)
    print("✅ 检查结果摘要")
    print("="*80)
    print(f"  检查成功: {result['checked']}")
    print(f"  页面总数: {result['page_total']}")
    print(f"  缓存总数: {result['cached_total']}")
    print(f"  数据匹配: {result['matched']}")
    print(f"  已更新: {result['updated']}")
    print(f"  消息: {result['message']}")
    print("="*80 + "\n")
    
    # 显示完整日志
    print(f"📋 收集了 {len(logs)} 条日志消息")
    
    return result


if __name__ == "__main__":
    try:
        result = test_startup_check()
        
        # 根据结果返回不同的退出码
        if result['checked']:
            sys.exit(0)
        else:
            sys.exit(1)
    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
