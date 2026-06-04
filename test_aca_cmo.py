#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文件 - 验证 ACA CMO 项目是否能被正确查询
使用账号/密码: schhs334/schhs334
"""

import sys
from pathlib import Path

# 添加 sms_app 到路径
sys.path.insert(0, str(Path(__file__).parent / "sms_app"))

from core.sms_handler import SMSHandler
import time


def test_aca_cmo_search():
    """测试搜索 ACA CMO 项目"""
    
    print("=" * 60)
    print("🧪 ACA CMO 项目搜索测试")
    print("=" * 60)
    print()
    
    # 1. 测试连接
    print("📍 第一步: 测试连接")
    print("-" * 60)
    handler = SMSHandler(headless=False)
    
    if handler.test_connection("schhs334", "schhs334"):
        print("✅ 连接测试通过")
    else:
        print("❌ 连接测试失败，请检查网络或凭证")
        return
    
    print()
    time.sleep(1)
    
    # 2. 登入系统
    print("📍 第二步: 登入系统")
    print("-" * 60)
    
    if handler.login("schhs334", "schhs334"):
        print("✅ 登入成功")
    else:
        print("❌ 登入失败")
        handler.close_driver()
        return
    
    print()
    time.sleep(2)
    
    # 3. 搜索 ACA CMO 项目
    print("📍 第三步: 搜索 ACA CMO 项目")
    print("-" * 60)
    
    projects = handler.search_projects("ACA CMO")
    
    if projects is None:
        print("❌ 搜索失败")
    elif len(projects) == 0:
        print("⚠️  未找到 ACA CMO 项目")
    else:
        print(f"✅ 找到 {len(projects)} 个项目:")
        print()
        for i, proj in enumerate(projects, 1):
            print(f"   项目 {i}:")
            for key, value in proj.items():
                print(f"      {key}: {value}")
            print()
    
    print()
    time.sleep(1)
    
    # 4. 搜索 ACA C（部分匹配）
    print("📍 第四步: 搜索 ACA C（部分匹配测试）")
    print("-" * 60)
    
    projects = handler.search_projects("ACA C")
    
    if projects is None:
        print("❌ 搜索失败")
    elif len(projects) == 0:
        print("⚠️  未找到包含 'ACA C' 的项目")
    else:
        print(f"✅ 找到 {len(projects)} 个项目:")
        print()
        for i, proj in enumerate(projects, 1):
            print(f"   项目 {i}:")
            for key, value in proj.items():
                print(f"      {key}: {value}")
            print()
    
    print()
    print("=" * 60)
    print("✅ 测试完成")
    print("=" * 60)
    
    # 关闭浏览器
    handler.close_driver()


if __name__ == "__main__":
    test_aca_cmo_search()
