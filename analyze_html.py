#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析 HTML 页面结构 - 找出搜索字段的正确名称
"""

import requests
import re

requests.packages.urllib3.disable_warnings()

LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
ITEM_SETTING_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"

session = requests.Session()
session.verify = False

# 登入
print("📍 登入系统...")
login_data = {
    'LoginForm[username]': 'schhs334',
    'LoginForm[password]': 'schhs334',
    'login-button': 'login'
}

response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
print("✓ 已登入\n")

# 访问项目设置页面
print("📍 访问项目设置页面...")
response = session.get(ITEM_SETTING_PAGE, timeout=10)
html = response.text

# 保存HTML以便分析
with open('itemsetting_page.html', 'w', encoding='utf-8') as f:
    f.write(html)
print("✓ 已保存 HTML 到 itemsetting_page.html\n")

# 查找搜索相关的 input 字段
print("🔍 搜索页面中的所有 input 字段：")
print("-" * 80)

# 查找所有 input 标签
inputs = re.findall(r'<input[^>]*>', html)
for i, inp in enumerate(inputs[:20], 1):  # 只显示前 20 个
    # 提取 name 和 placeholder
    name_match = re.search(r'name=["\']([^"\']+)["\']', inp)
    placeholder_match = re.search(r'placeholder=["\']([^"\']+)["\']', inp)
    type_match = re.search(r'type=["\']([^"\']+)["\']', inp)
    
    if name_match:
        name = name_match.group(1)
        placeholder = placeholder_match.group(1) if placeholder_match else ""
        input_type = type_match.group(1) if type_match else "text"
        
        print(f"[{i}] name={name}")
        print(f"    type={input_type}")
        if placeholder:
            print(f"    placeholder={placeholder}")
        print()

session.close()
