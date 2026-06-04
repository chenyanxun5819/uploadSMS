#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试改进的搜索方法
"""

import sys
sys.path.insert(0, './sms_app')

from core.sms_handler import SMSHandler

def test_search():
    """测试搜索功能"""
    print("=" * 90)
    print("🧪 测试改进的搜索方法 - ACA CMI")
    print("=" * 90)
    print()
    
    handler = SMSHandler(headless=False)
    
    # 先登入
    print("📍 第一步: 登入...")
    if not handler.login('schhs334', 'schhs334', timeout=15):
        print("❌ 登入失败")
        return
    
    print()
    
    # 搜索 ACA CMI
    print("📍 第二步: 搜索 ACA CMI（最新5条）...")
    print()
    
    projects = handler.search_projects_and_get_latest('ACA CMI', limit=5)
    
    print()
    print("=" * 90)
    
    if projects is None:
        print("❌ 搜索失败")
    elif len(projects) == 0:
        print("⚠️  未找到 ACA CMI 项目")
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
