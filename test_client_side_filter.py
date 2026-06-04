#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高效的项目搜索方法 - 客户端本地过滤
不需要 WebDriver，直接获取页面然后在本地过滤
"""

import requests
from html.parser import HTMLParser
import time

requests.packages.urllib3.disable_warnings()

class ProjectTableParser(HTMLParser):
    """项目表格解析器 - 提取所有表格行"""
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

print("=" * 90)
print("🧪 测试客户端本地过滤方法")
print("=" * 90)
print()

session = requests.Session()
session.verify = False

# 登入
print("📍 步骤1: 登入系统...", end="", flush=True)
login_data = {
    'LoginForm[username]': 'schhs334',
    'LoginForm[password]': 'schhs334',
    'login-button': 'login'
}
session.post(LOGIN_URL, data=login_data, timeout=10)
print(" ✅\n")

# 获取项目列表页面（不带任何搜索参数）
print("📍 步骤2: 获取完整的项目列表页面（第1页）...", end="", flush=True)
response = session.get(ITEM_SETTING_PAGE, timeout=10)
parser = ProjectTableParser()
parser.feed(response.text)
all_rows = parser.rows
print(f" ✅ ({len(all_rows)} 行)\n")

# 在本地过滤
print("📍 步骤3: 在本地过滤匹配 'ACA CMI' 的项目...\n")

filtered_projects = []
for row in all_rows:
    if len(row) >= 3:
        seq = row[0].strip()
        code = row[1].strip()
        name = row[2].strip()
        
        # 客户端过滤：项目代码以搜索条件开头
        if code.startswith("ACA CMI"):
            filtered_projects.append({
                '序号': seq,
                '项目代码': code,
                '项目名称': name
            })
            print(f"  ✓ 找到: {code} - {name[:40]}")

print()
print("=" * 90)
print(f"🎯 搜索结果: 共找到 {len(filtered_projects)} 个项目")
print("=" * 90)

if filtered_projects:
    print("\n项目列表（最新在前）：")
    print("-" * 90)
    print(f"{'序号':<6} {'项目代码':<20} {'项目名称':<60}")
    print("-" * 90)
    
    # 倒序排列
    filtered_projects = list(reversed(filtered_projects))
    for project in filtered_projects[:5]:  # 只显示前5条
        print(f"{project['序号']:<6} {project['项目代码']:<20} {project['项目名称']:<60}")
else:
    print("\n⚠️  未找到匹配的项目")

session.close()
