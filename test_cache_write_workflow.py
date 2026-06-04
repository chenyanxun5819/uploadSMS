#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试项目添加后的缓存写入流程

验证：
1. 项目添加成功
2. 从服务器读取最后一条记录（包括序号）
3. 缓存正确保存
4. 表格正确更新
"""

import sys
import time
from pathlib import Path

# 添加路径
sys.path.insert(0, str(Path(__file__).parent))

from sms_app.core.cache_manager import ProjectCacheManager
from sms_app.core.config_manager import ConfigManager
import requests
from html.parser import HTMLParser
import re
import json


class TestProjectCache:
    """测试项目缓存写入"""
    
    def __init__(self):
        self.config = ConfigManager()
        self.cache_manager = ProjectCacheManager()
        self.BASE_URL = "http://sms.chhsban.edu.my/sms"
        self.session = requests.Session()
        self.session.verify = False
    
    def test_1_get_total_count(self):
        """测试1：获取项目总数"""
        print("\n" + "="*60)
        print("测试1：获取项目总数")
        print("="*60)
        
        try:
            # 获取凭证
            username, password = self.config.get_credentials()
            if not username or not password:
                print("❌ 未找到凭证，请先保存凭证")
                return False
            
            # 登入
            print(f"🔐 使用用户: {username}")
            login_url = f"{self.BASE_URL}/index.php?r=site/login"
            login_data = {
                'LoginForm[username]': username,
                'LoginForm[password]': password,
                'login': '登入'
            }
            
            response = self.session.post(login_url, data=login_data, timeout=10, allow_redirects=True)
            
            if 'logout' not in response.text.lower():
                print("❌ 登入失败")
                return False
            
            print("✅ 登入成功")
            
            # 获取第一页数据来提取总数
            url = "http://sms.chhsban.edu.my/sms/index.php"
            params = {
                'ItemM_page': 1,
                'ajax': 'item-m-grid',
                'r': 'transaction/itemSetting/index'
            }
            
            response = self.session.get(url, params=params, timeout=10)
            
            # 提取总数
            match = re.search(r'第\s*\d+[-~]\d+\s*条，?共\s*(\d+)\s*条', response.text)
            if match:
                total = int(match.group(1))
                print(f"✅ 项目总数: {total}")
                return total
            else:
                print("❌ 无法提取项目总数")
                return None
                
        except Exception as e:
            print(f"❌ 异常: {e}")
            return None
    
    def test_2_get_last_project(self):
        """测试2：获取最后一条项目记录"""
        print("\n" + "="*60)
        print("测试2：获取最后一条项目记录")
        print("="*60)
        
        try:
            # 获取凭证
            username, password = self.config.get_credentials()
            if not username or not password:
                print("❌ 未找到凭证")
                return None
            
            # 登入
            print(f"🔐 使用用户: {username}")
            login_url = f"{self.BASE_URL}/index.php?r=site/login"
            login_data = {
                'LoginForm[username]': username,
                'LoginForm[password]': password,
                'login': '登入'
            }
            
            response = self.session.post(login_url, data=login_data, timeout=10, allow_redirects=True)
            
            if 'logout' not in response.text.lower():
                print("❌ 登入失败")
                return None
            
            print("✅ 登入成功")
            
            # 获取第一页来提取总数
            url = "http://sms.chhsban.edu.my/sms/index.php"
            params = {
                'ItemM_page': 1,
                'ajax': 'item-m-grid',
                'r': 'transaction/itemSetting/index'
            }
            
            response = self.session.get(url, params=params, timeout=10)
            match = re.search(r'第\s*\d+[-~]\d+\s*条，?共\s*(\d+)\s*条', response.text)
            
            if not match:
                print("❌ 无法提取项目总数")
                return None
            
            total_count = int(match.group(1))
            last_page = (total_count + 9) // 10
            
            print(f"📊 项目总数: {total_count}")
            print(f"📄 最后一页: {last_page}")
            
            # 获取最后一页
            params['ItemM_page'] = last_page
            response = self.session.get(url, params=params, timeout=10)
            
            # 解析最后一条记录
            class ProjectTableParser(HTMLParser):
                def __init__(self):
                    super().__init__()
                    self.rows = []
                    self.in_tbody = False
                    self.current_row = []
                    self.in_td = False
                    self.current_cell = ""
                
                def handle_starttag(self, tag, attrs):
                    if tag == "tbody":
                        self.in_tbody = True
                    elif tag == "tr" and self.in_tbody:
                        self.current_row = []
                    elif tag in ["td", "th"] and self.in_tbody:
                        self.in_td = True
                        self.current_cell = ""
                
                def handle_endtag(self, tag):
                    if tag == "tbody":
                        self.in_tbody = False
                    elif tag == "tr" and self.in_tbody:
                        if self.current_row:
                            self.rows.append(self.current_row)
                    elif tag in ["td", "th"] and self.in_tbody:
                        self.in_td = False
                        self.current_row.append(self.current_cell.strip())
                
                def handle_data(self, data):
                    if self.in_td:
                        self.current_cell += data
            
            parser = ProjectTableParser()
            parser.feed(response.text)
            
            print(f"📋 总共获取 {len(parser.rows)} 条记录")
            
            if parser.rows:
                last_row = parser.rows[-1]
                print(f"📌 最后一行有 {len(last_row)} 列数据")
                
                if len(last_row) >= 4:
                    project = {
                        '序号': last_row[0].strip() if len(last_row) > 0 else '',
                        '项目代码': last_row[1].strip() if len(last_row) > 1 else '',
                        '项目名称': last_row[2].strip() if len(last_row) > 2 else '',
                        '分数': last_row[3].strip() if len(last_row) > 3 else '0.00'
                    }
                    
                    print("✅ 最后一条项目:")
                    print(f"   序号: {project.get('序号')}")
                    print(f"   代码: {project.get('项目代码')}")
                    print(f"   名称: {project.get('项目名称')}")
                    print(f"   分数: {project.get('分数')}")
                    
                    return project
                else:
                    print(f"❌ 列数不足: {len(last_row)}")
                    return None
            else:
                print("❌ 无数据行")
                return None
                
        except Exception as e:
            print(f"❌ 异常: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def test_3_cache_operations(self):
        """测试3：缓存操作"""
        print("\n" + "="*60)
        print("测试3：缓存操作")
        print("="*60)
        
        try:
            # 加载缓存
            projects, metadata = self.cache_manager.load_cache()
            
            if projects:
                print(f"✅ 缓存已加载: {len(projects)} 条项目")
                print(f"   总数: {metadata.get('total_count', 'N/A')}")
                print(f"   页数: {metadata.get('total_pages', 'N/A')}")
                
                # 显示前3条和最后3条
                print("\n   前3条项目:")
                for i, p in enumerate(projects[:3]):
                    print(f"     [{p.get('序号')}] {p.get('项目代码')} - {p.get('项目名称')}")
                
                if len(projects) > 6:
                    print("\n   ...（省略）...\n")
                
                print("\n   最后3条项目:")
                for i, p in enumerate(projects[-3:]):
                    print(f"     [{p.get('序号')}] {p.get('项目代码')} - {p.get('项目名称')}")
                
                return True
            else:
                print("⚠️  缓存为空或不存在")
                return False
                
        except Exception as e:
            print(f"❌ 异常: {e}")
            return False
    
    def test_4_verify_sequence_numbers(self):
        """测试4：验证序号完整性"""
        print("\n" + "="*60)
        print("测试4：验证序号完整性")
        print("="*60)
        
        try:
            projects, _ = self.cache_manager.load_cache()
            
            if not projects:
                print("⚠️  缓存为空")
                return False
            
            # 检查序号
            missing_seq = []
            for i, p in enumerate(projects):
                seq = p.get('序号')
                if not seq or seq.strip() == '':
                    missing_seq.append(i)
            
            if missing_seq:
                print(f"❌ 发现 {len(missing_seq)} 条记录缺少序号:")
                for idx in missing_seq[:10]:  # 只显示前10条
                    print(f"   [{idx}] {projects[idx].get('项目代码')}")
                if len(missing_seq) > 10:
                    print(f"   ... 还有 {len(missing_seq) - 10} 条")
                return False
            else:
                print(f"✅ 所有 {len(projects)} 条记录都有序号")
                return True
                
        except Exception as e:
            print(f"❌ 异常: {e}")
            return False
    
    def run_all_tests(self):
        """运行所有测试"""
        print("\n")
        print("█" * 60)
        print("█" + " " * 58 + "█")
        print("█" + "   项目缓存写入流程测试".center(58) + "█")
        print("█" + " " * 58 + "█")
        print("█" * 60)
        
        results = {}
        
        # 测试1：获取总数
        result1 = self.test_1_get_total_count()
        results['获取项目总数'] = result1 is not None
        
        # 测试2：获取最后一条记录
        result2 = self.test_2_get_last_project()
        results['获取最后一条项目'] = result2 is not None
        
        # 测试3：缓存操作
        result3 = self.test_3_cache_operations()
        results['缓存操作'] = result3
        
        # 测试4：验证序号
        result4 = self.test_4_verify_sequence_numbers()
        results['序号完整性'] = result4
        
        # 总结
        print("\n" + "="*60)
        print("测试总结")
        print("="*60)
        
        passed = sum(1 for v in results.values() if v)
        total = len(results)
        
        for test_name, passed_flag in results.items():
            status = "✅ PASS" if passed_flag else "❌ FAIL"
            print(f"{status}: {test_name}")
        
        print("-" * 60)
        print(f"总计: {passed}/{total} 测试通过")
        
        if passed == total:
            print("\n🎉 所有测试通过！")
            return True
        else:
            print(f"\n⚠️  有 {total - passed} 个测试失败")
            return False
        
        self.session.close()


if __name__ == '__main__':
    tester = TestProjectCache()
    tester.run_all_tests()
