#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速验证脚本 - 检查实现是否正确完成

使用方式:
    python verify_implementation.py
"""

import sys
import os
from pathlib import Path

def check_file_exists(path, name):
    """检查文件是否存在"""
    if os.path.exists(path):
        print(f"✅ {name}: {path}")
        return True
    else:
        print(f"❌ {name}: {path} (不存在)")
        return False

def check_string_in_file(path, search_string, name):
    """检查文件中是否包含某个字符串"""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
            if search_string in content:
                print(f"✅ {name}")
                return True
            else:
                print(f"❌ {name}")
                return False
    except Exception as e:
        print(f"❌ {name}: 读取文件失败 - {e}")
        return False

def check_class_exists(path, class_name):
    """检查类是否存在"""
    return check_string_in_file(path, f"class {class_name}", f"类 {class_name}")

def check_method_exists(path, method_name):
    """检查方法是否存在"""
    return check_string_in_file(path, f"def {method_name}", f"方法 {method_name}")

def main():
    print("\n" + "="*70)
    print("项目添加缓存同步 - 实现验证".center(70))
    print("="*70 + "\n")
    
    # 文件路径
    project_input_page = "sms_app/ui/pages/project_input_page.py"
    
    if not os.path.exists(project_input_page):
        print(f"❌ 找不到核心文件: {project_input_page}")
        print("   请在项目根目录运行此脚本")
        return False
    
    results = []
    
    # 1. 检查文件存在
    print("📋 步骤 1: 检查文件完整性")
    print("-" * 70)
    results.append(check_file_exists(project_input_page, "项目输入页面"))
    results.append(check_file_exists("sms_app/core/cache_manager.py", "缓存管理器"))
    results.append(check_file_exists("sms_app/core/config_manager.py", "配置管理器"))
    
    # 2. 检查类是否存在
    print("\n📋 步骤 2: 检查新增类")
    print("-" * 70)
    results.append(check_class_exists(project_input_page, "FetchLastProjectThread"))
    
    # 3. 检查新增方法
    print("\n📋 步骤 3: 检查新增方法")
    print("-" * 70)
    results.append(check_method_exists(project_input_page, "_on_fetch_last_project_finished"))
    results.append(check_method_exists(project_input_page, "_load_projects_from_cache"))
    
    # 4. 检查 FetchLastProjectThread 中的方法
    print("\n📋 步骤 4: 检查 FetchLastProjectThread 的方法")
    print("-" * 70)
    results.append(check_string_in_file(project_input_page, "def run(self):", "run() 方法"))
    results.append(check_string_in_file(project_input_page, "def _login(self, username", "_login() 方法"))
    results.append(check_string_in_file(project_input_page, "def _get_total_count(self)", "_get_total_count() 方法"))
    results.append(check_string_in_file(project_input_page, "def _fetch_last_project(self)", "_fetch_last_project() 方法"))
    
    # 5. 检查导入
    print("\n📋 步骤 5: 检查必要的导入")
    print("-" * 70)
    results.append(check_string_in_file(project_input_page, "import re", "import re"))
    results.append(check_string_in_file(project_input_page, "from html.parser import HTMLParser", "HTMLParser 导入"))
    results.append(check_string_in_file(project_input_page, "import requests", "requests 导入"))
    
    # 6. 检查信号定义
    print("\n📋 步骤 6: 检查信号定义")
    print("-" * 70)
    results.append(check_string_in_file(project_input_page, "fetch_finished = pyqtSignal", "fetch_finished 信号"))
    
    # 7. 检查正则表达式
    print("\n📋 步骤 7: 检查关键代码")
    print("-" * 70)
    results.append(check_string_in_file(project_input_page, r"r'第", "项目总数正则表达式"))
    results.append(check_string_in_file(project_input_page, "ItemM_page", "AJAX 参数"))
    results.append(check_string_in_file(project_input_page, "item-m-grid", "AJAX 端点参数"))
    
    # 8. 检查 _on_add_finished 的修改
    print("\n📋 步骤 8: 检查 _on_add_finished 方法的修改")
    print("-" * 70)
    results.append(check_string_in_file(project_input_page, "FetchLastProjectThread(", "启动 FetchLastProjectThread"))
    results.append(check_string_in_file(project_input_page, "fetch_finished.connect", "连接 fetch_finished 信号"))
    
    # 9. 检查缓存操作
    print("\n📋 步骤 9: 检查缓存操作")
    print("-" * 70)
    results.append(check_string_in_file(project_input_page, "ProjectCacheManager()", "缓存管理器初始化"))
    results.append(check_string_in_file(project_input_page, ".load_cache()", "加载缓存方法"))
    results.append(check_string_in_file(project_input_page, ".save_cache(", "保存缓存方法"))
    
    # 10. 检查文档
    print("\n📋 步骤 10: 检查文档完整性")
    print("-" * 70)
    results.append(check_file_exists("CACHE_WRITE_OPTIMIZATION.md", "优化文档"))
    results.append(check_file_exists("PROJECT_ADD_QUICK_REF.md", "快速参考"))
    results.append(check_file_exists("IMPLEMENTATION_COMPLETE.md", "实现完成报告"))
    
    # 总结
    print("\n" + "="*70)
    print("验证结果总结".center(70))
    print("="*70 + "\n")
    
    passed = sum(results)
    total = len(results)
    
    print(f"✅ 通过项目: {passed}")
    print(f"❌ 失败项目: {total - passed}")
    print(f"📊 通过率: {passed}/{total} ({100*passed//total}%)")
    
    if passed == total:
        print("\n" + "🎉 " * 20)
        print("所有验证项目都通过了！".center(70))
        print("实现已准备好进行端到端测试。".center(70))
        print("🎉 " * 20)
        return True
    else:
        print(f"\n⚠️  还有 {total - passed} 个项目需要修复")
        return False
    
    print("\n" + "="*70 + "\n")

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
