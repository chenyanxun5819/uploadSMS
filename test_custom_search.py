#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试自定义项目搜索功能
"""

import requests
from html.parser import HTMLParser
import time

requests.packages.urllib3.disable_warnings()

LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
ITEM_SETTING_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"

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

def search_projects(search_prefix: str):
    """搜索项目"""
    
    session = requests.Session()
    session.verify = False
    
    print("="*70)
    print(f"🔍 搜索项目: '{search_prefix}'")
    print("="*70)
    
    # 登入
    print(f"  📍 登入系统...", end="", flush=True)
    login_data = {
        'LoginForm[username]': 'schhs334',
        'LoginForm[password]': 'schhs334',
        'login-button': 'login'
    }
    session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
    print(" ✅")
    
    time.sleep(0.5)
    
    # 获取所有项目
    print(f"  📍 获取项目列表...", end="", flush=True)
    all_projects = []
    
    for page in range(1, 101):
        try:
            response = session.get(ITEM_SETTING_PAGE, params={'page': page}, timeout=10)
            parser = ProjectTableParser()
            parser.feed(response.text)
            
            if len(parser.rows) > 0:
                for row in parser.rows:
                    if len(row) >= 3:
                        project = {
                            '序号': len(all_projects) + 1,
                            '项目代码': row[1].strip(),
                            '项目名称': row[2].strip()
                        }
                        all_projects.append(project)
            else:
                break
        except Exception as e:
            break
        
        if page % 10 == 0:
            print(f"\r  📍 获取项目列表...({page} 页，{len(all_projects)} 项)", end="", flush=True)
    
    print(f"\r  📍 获取项目列表... ✅ ({len(all_projects)} 项)", end="")
    print()
    
    # 本地过滤
    print(f"  📍 过滤项目 '{search_prefix}'...", end="", flush=True)
    
    filtered_projects = []
    for project in all_projects:
        code = project.get('项目代码', '')
        # 搜索项目代码中包含搜索前缀（不区分大小写）
        if search_prefix.upper() in code.upper():
            filtered_projects.append(project)
    
    print(f" ✅ ({len(filtered_projects)} 项)")
    
    # 倒序
    filtered_projects = list(reversed(filtered_projects))
    
    session.close()
    
    # 显示结果
    print()
    print("="*70)
    if len(filtered_projects) > 0:
        print(f"✅ 共找到 {len(filtered_projects)} 个项目")
        print("="*70)
        
        for idx, project in enumerate(filtered_projects, 1):
            print(f"  {idx}. {project['项目代码']:<20} | {project['项目名称']}")
    else:
        print(f"⚠️  未找到匹配 '{search_prefix}' 的项目")
        print("="*70)
        print()
        print("📋 系统中现有的项目代码：")
        # 显示所有项目代码
        unique_codes = sorted(set([p['项目代码'] for p in all_projects]))
        for code in unique_codes:
            count = sum(1 for p in all_projects if p['项目代码'] == code)
            print(f"  - {code:<20} ({count} 个)")
    
    print()

if __name__ == "__main__":
    # 测试不同的搜索前缀
    test_searches = [
        "ACA CMI",      # 你想要的
        "ACA CMO",      # 应该能找到
        "CCD PO",       # 应该能找到
        "PE CMO",       # 应该能找到
    ]
    
    for search_term in test_searches:
        search_projects(search_term)
        print()
