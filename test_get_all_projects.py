#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
获取足够多的页数，然后本地过滤 ACA CMI
"""

import requests
from html.parser import HTMLParser
import time

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

print("=" * 90)
print("🧪 获取所有项目，本地过滤 ACA CMI")
print("=" * 90)
print()

session = requests.Session()
session.verify = False

# 登入
print("📍 登入系统...", end="", flush=True)
login_data = {
    'LoginForm[username]': 'schhs334',
    'LoginForm[password]': 'schhs334',
    'login-button': 'login'
}
session.post(LOGIN_URL, data=login_data, timeout=10)
print(" ✅\n")

# 获取足够多的页（500+ 页，应该能覆盖所有项目）
print("📍 获取所有项目列表...\n")

all_rows = []
page = 1
max_pages = 500  # 获取很多页
consecutive_empty_pages = 0  # 连续空页计数

while page <= max_pages:
    if page % 50 == 1:  # 每 50 页打印一次进度
        print(f"  ⏳ 获取第 {page}-{page+49} 页...", end="", flush=True)
    
    try:
        response = session.get(ITEM_SETTING_PAGE, params={'page': page}, timeout=10)
        parser = ProjectTableParser()
        parser.feed(response.text)
        
        if len(parser.rows) == 0:
            consecutive_empty_pages += 1
            # 如果连续 5 页都是空的，说明已经到了最后
            if consecutive_empty_pages >= 5:
                if page % 50 != 1:
                    print()
                print(f"  ✓ 已到达最后一页（第 {page-5} 页）")
                break
        else:
            consecutive_empty_pages = 0
            all_rows.extend(parser.rows)
        
        if page % 50 == 0:
            print(f" ✓")
        
        page += 1
        time.sleep(0.05)  # 稍微延迟一下
        
    except Exception as e:
        print(f" ❌ 错误: {e}")
        break

print(f"\n✅ 共获取 {len(all_rows)} 行项目数据\n")

# 在本地过滤 ACA CMI
print("🔍 在本地过滤 'ACA CMI' 的项目...\n")

search_code = "ACA CMI"
filtered_projects = []

for row in all_rows:
    if len(row) >= 3:
        seq = row[0].strip()
        code = row[1].strip()
        name = row[2].strip()
        
        # 过滤：项目代码以搜索条件开头
        if code.startswith(search_code):
            filtered_projects.append({
                '序号': seq,
                '项目代码': code,
                '项目名称': name
            })

print()
print("=" * 90)
print(f"🎯 搜索结果: 共找到 {len(filtered_projects)} 个 ACA CMI 项目")
print("=" * 90)

if filtered_projects:
    print("\n项目列表（最新在前）：")
    print("-" * 90)
    print(f"{'序号':<6} {'项目代码':<20} {'项目名称':<60}")
    print("-" * 90)
    
    # 倒序排列（最新的在前）
    filtered_projects = list(reversed(filtered_projects))
    for project in filtered_projects[:5]:  # 只显示前5条
        print(f"{project['序号']:<6} {project['项目代码']:<20} {project['项目名称']:<60}")
    
    if len(filtered_projects) > 5:
        print(f"... 还有 {len(filtered_projects) - 5} 个项目")
else:
    print(f"\n⚠️  未找到 '{search_code}' 的项目")
    print("\n📝 页面中找到的项目代码示例（前10个）：")
    for row in all_rows[:10]:
        if len(row) >= 2:
            print(f"  - {row[1].strip()}")

session.close()
