#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试搜索参数 - 检查是否正确传递搜索条件
"""

import requests
from html.parser import HTMLParser

class ProjectTableParser(HTMLParser):
    """项目表格解析器"""
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

# 禁用 SSL 警告
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
print(f"✓ 已登入\n")

# 测试不同的搜索参数
test_cases = [
    {
        'name': '方法1: MarkItem[code]',
        'params': {'MarkItem[code]': 'ACA CMI', 'page': 1}
    },
    {
        'name': '方法2: mark_item (snake_case)',
        'params': {'mark_item': 'ACA CMI', 'page': 1}
    },
    {
        'name': '方法3: StudentPerformance[mark_item]',
        'params': {'StudentPerformance[mark_item]': 'ACA CMI', 'page': 1}
    },
]

for test_case in test_cases:
    print(f"🧪 测试: {test_case['name']}")
    print(f"   参数: {test_case['params']}")
    
    try:
        response = session.get(ITEM_SETTING_PAGE, params=test_case['params'], timeout=10)
        
        parser = ProjectTableParser()
        parser.feed(response.text)
        rows = parser.rows
        
        print(f"   找到: {len(rows)} 行数据")
        
        if rows and len(rows) > 0:
            print(f"   第一行项目代码: {rows[0][1] if len(rows[0]) > 1 else 'N/A'}")
            print(f"   第一行项目名称: {rows[0][2] if len(rows[0]) > 2 else 'N/A'}")
        
        print()
    except Exception as e:
        print(f"   ❌ 错误: {e}\n")

session.close()
