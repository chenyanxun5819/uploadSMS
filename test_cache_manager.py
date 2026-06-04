#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试缓存管理器 - 下载所有 2414 条项目数据
"""

from sms_app.core.cache_manager import ProjectCacheManager

def main():
    manager = ProjectCacheManager()
    
    print("\n" + "="*70)
    print("🧪 测试缓存管理器")
    print("="*70 + "\n")
    
    # 检查是否有缓存
    if manager.has_cache():
        print("✅ 已存在缓存")
        info = manager.get_cache_info()
        print(f"  项目数：{info['project_count']}")
        print(f"  更新时间：{info['last_updated']}")
        print(f"  缓存位置：{info['cache_dir']}")
        
        # 加载缓存进行搜索测试
        projects, metadata = manager.load_cache()
        
        # 测试搜索
        test_searches = [
            "ACA CMI",
            "CCD CMO",
            "PE CMO"
        ]
        
        print("\n" + "="*70)
        print("🔍 测试搜索")
        print("="*70 + "\n")
        
        for search_term in test_searches:
            filtered = [p for p in projects if search_term.upper() in p['项目代码'].upper()]
            print(f"搜索 '{search_term}': 找到 {len(filtered)} 项")
            if filtered:
                # 显示前 3 项
                for i, p in enumerate(filtered[:3], 1):
                    print(f"  {i}. {p['项目代码']} - {p['项目名称'][:50]}")
                if len(filtered) > 3:
                    print(f"  ... 还有 {len(filtered)-3} 项")
        
    else:
        print("❌ 未找到缓存，需要下载")
        
        # 请求凭证
        username = "schhs334"
        password = "schhs334"
        
        print(f"\n使用凭证：{username}")
        
        # 开始下载
        result = manager.download_all_projects(username, password)
        
        if result['success']:
            projects = result['projects']
            metadata = result['metadata']
            
            # 保存到缓存
            manager.save_cache(projects, metadata)
            
            print("\n📊 下载统计：")
            print(f"  总项目数：{len(projects)}")
            print(f"  总页数：{metadata['total_pages']}")
            print(f"  首条项目：{metadata['first_project_id']}")
            print(f"  末条项目：{metadata['last_project_id']}")
            
            # 测试搜索
            print("\n" + "="*70)
            print("🔍 测试搜索")
            print("="*70 + "\n")
            
            test_searches = [
                "ACA CMI",
                "CCD CMO",
                "PE CMO",
                "CMI"
            ]
            
            for search_term in test_searches:
                filtered = [p for p in projects if search_term.upper() in p['项目代码'].upper()]
                print(f"搜索 '{search_term}': 找到 {len(filtered)} 项")
                if filtered:
                    # 显示前 3 项和最后 3 项
                    print(f"  最新 3 项：")
                    for i, p in enumerate(filtered[-3:], 1):
                        print(f"    {i}. {p['项目代码']} - {p['项目名称'][:40]}")
        else:
            print(f"\n❌ 下载失败：{result.get('error', '未知错误')}")

if __name__ == "__main__":
    main()
