#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AJAX 端点连接测试
验证从 AJAX 端点获取数据是否正常
"""

import requests
import re
import time

requests.packages.urllib3.disable_warnings()


def test_ajax_endpoint():
    """测试 AJAX 端点连接"""
    
    print("\n" + "="*80)
    print("🧪 AJAX 端点连接测试")
    print("="*80 + "\n")
    
    # 测试参数
    USERNAME = "input_your_username_here"
    PASSWORD = "input_your_password_here"
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    AJAX_URL = "http://sms.chhsban.edu.my/sms/index.php"
    
    session = requests.Session()
    session.verify = False
    
    try:
        # 第1步：登入
        print("📍 步骤 1: 登入系统...")
        if not USERNAME.startswith("input_"):
            login_data = {
                'LoginForm[username]': USERNAME,
                'LoginForm[password]': PASSWORD,
                'login-button': 'login'
            }
            response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
            print(f"  ✅ 登入成功 (状态码: {response.status_code})")
        else:
            print("  ⚠️  未配置凭证，跳过登入")
            return
        
        time.sleep(0.5)
        
        # 第2步：测试 AJAX 端点
        print("\n📍 步骤 2: 测试 AJAX 端点...")
        params = {
            'ItemM_page': 1,
            'ajax': 'item-m-grid',
            'r': 'transaction/itemSetting/index'
        }
        
        print(f"  📥 请求 URL: {AJAX_URL}")
        print(f"  📥 参数: {params}")
        
        response = session.get(AJAX_URL, params=params, timeout=10)
        print(f"  ✅ 连接成功 (状态码: {response.status_code})")
        
        # 第3步：提取总数
        print("\n📍 步骤 3: 从响应提取总数...")
        match = re.search(r'第\s*\d+[-~]\d+\s*条，?共\s*(\d+)\s*条', response.text)
        if match:
            total = int(match.group(1))
            print(f"  ✅ 成功提取总数: {total} 条")
        else:
            print(f"  ❌ 无法提取总数")
            print(f"  📋 响应长度: {len(response.text)} 字符")
            print(f"  📋 响应前 500 字符:")
            print(f"     {response.text[:500]}")
        
        # 第4步：提取项目数据
        print("\n📍 步骤 4: 从响应提取项目数据...")
        from html.parser import HTMLParser
        
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
        
        print(f"  ✅ 成功提取 {len(parser.rows)} 行项目数据")
        
        if parser.rows:
            print(f"\n  📋 前 3 行项目数据:")
            for i, row in enumerate(parser.rows[:3]):
                print(f"     行 {i+1}: {row}")
        
        session.close()
        print("\n" + "="*80)
        print("✅ AJAX 端点测试完成")
        print("="*80 + "\n")
        
    except Exception as e:
        print(f"  ❌ 测试失败: {type(e).__name__}: {e}")
        session.close()
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("\n" + "="*80)
    print("📝 注意：请在代码中填入您的用户名和密码")
    print("   找到这两行并替换:")
    print('   USERNAME = "input_your_username_here"')
    print('   PASSWORD = "input_your_password_here"')
    print("="*80)
    
    test_ajax_endpoint()
