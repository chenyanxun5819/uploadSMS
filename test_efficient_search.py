#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试高效搜索方法 - 不需要 WebDriver
"""

import sys
sys.path.insert(0, './sms_app')

from core.sms_handler import SMSHandler

def test_search():
    """测试搜索功能"""
    print("=" * 90)
    print("🧪 测试高效项目搜索（最后一页的最后5条记录）")
    print("=" * 90)
    print()
    
    # 创建处理器（不需要 WebDriver）
    handler = SMSHandler(headless=False)
    
    # 搜索 ACA CMO 的最后5条项目
    prefix_code = "ACA CMO"
    username = "schhs334"
    password = "schhs334"
    
    print(f"搜索前置编码: {prefix_code}")
    print(f"用户名: {username}")
    print()
    
    projects = handler.search_projects_efficient(
        username, 
        password, 
        prefix_code,
        limit=5
    )
    
    print()
    print("=" * 90)
    
    if projects is None:
        print("❌ 搜索失败")
    elif len(projects) == 0:
        print("⚠️  未找到项目")
    else:
        print(f"✅ 搜索成功！找到 {len(projects)} 个项目")
        print()
        print("项目列表（最新在前）：")
        print("-" * 90)
        print(f"{'序号':<6} {'项目代码':<20} {'项目名称':<60}")
        print("-" * 90)
        
        for idx, project in enumerate(projects, 1):
            seq = project.get('序号', '')
            code = project.get('项目代码', '')
            name = project.get('项目名称', '')
            print(f"{seq:<6} {code:<20} {name:<60}")
    
    print("=" * 90)

if __name__ == '__main__':
    test_search()
