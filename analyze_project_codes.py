#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析所有项目代码的格式
"""

import requests
from html.parser import HTMLParser
from collections import defaultdict

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

# 获取前 100 页
print("📍 获取前 100 页...")
all_rows = []

for page in range(1, 101):
    if page % 20 == 1:
        print(f"  获取第 {page}-{page+19} 页...", end="", flush=True)
    
    response = session.get(ITEM_SETTING_PAGE, params={'page': page}, timeout=10)
    parser = ProjectTableParser()
    parser.feed(response.text)
    
    if len(parser.rows) > 0:
        all_rows.extend(parser.rows)
    
    if page % 20 == 0:
        print(" ✓")

print(f"\n✅ 共获取 {len(all_rows)} 行\n")

# 分析项目代码
print("📊 分析项目代码...\n")

# 按前缀分类
prefix_count = defaultdict(int)
all_codes = []

for row in all_rows:
    if len(row) >= 2:
        code = row[1].strip()
        all_codes.append(code)
        
        # 提取前缀（通常是 "ACA", "CCD", "PE" 等）
        parts = code.split()
        if len(parts) >= 2:
            prefix = f"{parts[0]} {parts[1]}"  # 如 "ACA CMO", "CCD PO" 等
        else:
            prefix = parts[0]
        
        prefix_count[prefix] += 1

print("项目代码分布：")
print("-" * 50)

for prefix in sorted(prefix_count.keys()):
    count = prefix_count[prefix]
    print(f"  {prefix:<15} : {count:>5} 个项目")

print()
print("所有不同的项目代码（去重）：")
print("-" * 50)

unique_codes = sorted(set(all_codes))
aca_codes = [code for code in unique_codes if code.startswith('ACA')]

print(f"\n找到 {len(aca_codes)} 个 ACA 开头的项目代码：")
for code in aca_codes[:20]:  # 显示前 20 个
    print(f"  - {code}")

if len(aca_codes) > 20:
    print(f"  ... 还有 {len(aca_codes) - 20} 个")

session.close()
