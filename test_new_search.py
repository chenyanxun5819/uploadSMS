#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试搜索功能 - ACA CMO
"""

import sys
sys.path.insert(0, './sms_app')

from core.sms_handler import SMSHandler

def test_search():
    """测试搜索功能"""
    print("=" * 90)
    print("🧪 测试搜索功能 - ACA CMO（正确的活动代码）")
    print("=" * 90)
    print()
    
    handler = SMSHandler(headless=False)
    
    projects = handler.search_projects_and_get_latest('ACA CMO', limit=5, username='schhs334', password='schhs334')
    
    print()
    print("=" * 90)
    
    if projects is None:
        print("❌ 搜索失败")
    elif len(projects) == 0:
        print("⚠️  未找到 ACA CMO 项目")
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
    
    handler.close_driver()

if __name__ == '__main__':
    test_search()
