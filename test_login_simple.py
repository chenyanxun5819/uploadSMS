#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS 登入测试 - 使用 requests 库
"""

import requests
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import json
import time

# 禁用 SSL 警告
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

def test_login_simple(username, password):
    """测试 SMS 登入"""
    print("=" * 60)
    print(f"🧪 SMS 登入测试（简化版）")
    print("=" * 60)
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    
    try:
        # 创建会话
        print("\n[1] 创建 HTTP 会话...")
        session = requests.Session()
        session.verify = False
        print("✓ 会话已创建")
        
        # 先获取登入页面（获取 CSRF token 等）
        print(f"\n[2] 获取登入页面...")
        response = session.get(LOGIN_URL, timeout=10)
        print(f"✓ 状态码: {response.status_code}")
        print(f"✓ 页面大小: {len(response.text)} 字节")
        
        # 提交登入表单
        print(f"\n[3] 提交登入表单...")
        print(f"   帐号: {username}")
        print(f"   密码: {'*' * len(password)}")
        
        login_data = {
            'LoginForm[username]': username,
            'LoginForm[password]': password,
            'login-button': 'login'
        }
        
        response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
        print(f"✓ 提交完成")
        print(f"✓ 状态码: {response.status_code}")
        print(f"✓ 最终 URL: {response.url}")
        
        # 检查是否登入成功
        print(f"\n[4] 检查登入结果...")
        if 'login' not in response.url.lower():
            print("\n" + "=" * 60)
            print("✅ 登入成功！")
            print(f"   已重定向到: {response.url}")
            print("=" * 60)
            return True
        else:
            print("\n" + "=" * 60)
            print("❌ 登入失败 - 仍在登入页面")
            print(f"   URL: {response.url}")
            print("=" * 60)
            
            # 检查错误信息
            if '错误' in response.text or 'error' in response.text.lower():
                print("\n⚠️  页面包含错误信息")
            
            return False
            
    except Exception as e:
        print(f"\n❌ 异常发生: {e}")
        print("=" * 60)
        return False

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        print("使用方法: python test_login_simple.py <帐号> <密码>")
        print("示例: python test_login_simple.py schhs334 schhs334")
        sys.exit(1)
    
    username = sys.argv[1]
    password = sys.argv[2]
    
    result = test_login_simple(username, password)
    sys.exit(0 if result else 1)
