#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试搜索参数 - 找出正确的搜索方式
"""

import requests
from html.parser import HTMLParser

requests.packages.urllib3.disable_warnings()

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
session.post(LOGIN_URL, data=login_data, timeout=10)
print("✅ 已登入\n")

# 测试不同的搜索参数方式
test_cases = [
    {
        'name': '方法1: GET 参数 - ItemM[code]',
        'method': 'GET',
        'params': {'ItemM[code]': 'ACA CMI', 'page': 1},
        'data': None
    },
    {
        'name': '方法2: POST 数据 - ItemM[code]',
        'method': 'POST',
        'params': {'page': 1},
        'data': {'ItemM[code]': 'ACA CMI'}
    },
    {
        'name': '方法3: GET 参数 - code',
        'method': 'GET',
        'params': {'code': 'ACA CMI', 'page': 1},
        'data': None
    },
    {
        'name': '方法4: 获取更多页 (获取第 50 页)',
        'method': 'GET',
        'params': {'page': 50},
        'data': None
    },
]

for test_case in test_cases:
    print(f"🧪 测试: {test_case['name']}")
    
    try:
        if test_case['method'] == 'GET':
            response = session.get(ITEM_SETTING_PAGE, params=test_case['params'], timeout=10)
        else:
            response = session.post(ITEM_SETTING_PAGE, params=test_case['params'], data=test_case['data'], timeout=10)
        
        parser = ProjectTableParser()
        parser.feed(response.text)
        rows = parser.rows
        
        print(f"   结果: {len(rows)} 行数据")
        
        if rows:
            # 显示前 3 行
            for i, row in enumerate(rows[:3], 1):
                if len(row) >= 2:
                    code = row[1].strip()[:20]
                    name = row[2].strip()[:40] if len(row) > 2 else ""
                    print(f"     [{i}] {code} - {name}")
        
        # 检查是否有 ACA CMI
        has_aca_cmi = False
        for row in rows:
            if len(row) >= 2 and row[1].strip().startswith('ACA CMI'):
                has_aca_cmi = True
                break
        
        if has_aca_cmi:
            print(f"   ✅ 找到 ACA CMI 项目！")
        else:
            print(f"   ❌ 未找到 ACA CMI 项目")
        
        print()
    except Exception as e:
        print(f"   ❌ 错误: {e}\n")

session.close()
