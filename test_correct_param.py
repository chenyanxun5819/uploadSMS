#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用正确的参数名测试搜索
"""

import requests
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
session.post(LOGIN_URL, data=login_data, timeout=10)
print("✓ 已登入\n")

# 测试搜索 ACA CMI
print("🧪 搜索 ACA CMI...")
search_params = {
    'ItemM[mark_item]': 'ACA CMI',
    'page': 1
}

response = session.get(ITEM_SETTING_PAGE, params=search_params, timeout=10)
parser = ProjectTableParser()
parser.feed(response.text)
rows = parser.rows

print(f"✓ 找到 {len(rows)} 行数据\n")

if rows:
    print("项目列表（前5条）：")
    print("-" * 80)
    print(f"{'序号':<6} {'项目代码':<20} {'项目名称':<50}")
    print("-" * 80)
    
    for row in rows[:5]:
        if len(row) >= 3:
            seq = row[0].strip()
            code = row[1].strip()
            name = row[2].strip()
            print(f"{seq:<6} {code:<20} {name:<50}")
else:
    print("❌ 未找到任何数据")

session.close()
