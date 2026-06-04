#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证 AJAX 分页 URL - 检查是否能获取真实不同的数据
"""

import requests
from html.parser import HTMLParser

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

def test_ajax_pagination():
    session = requests.Session()
    session.verify = False
    
    print("="*70)
    print("🧪 测试 AJAX 分页 URL")
    print("="*70 + "\n")
    
    # 登入
    print("📍 登入系统...", end="", flush=True)
    login_data = {
        'LoginForm[username]': 'schhs334',
        'LoginForm[password]': 'schhs334',
        'login-button': 'login'
    }
    session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
    print(" ✅\n")
    
    # 测试多个页面
    pages_to_test = [1, 2, 3, 150, 151, 152, 240, 241, 242]
    
    print("📊 对比不同页码的数据：\n")
    
    all_projects_by_page = {}
    
    for page in pages_to_test:
        # 构建 AJAX URL
        url = "http://sms.chhsban.edu.my/sms/index.php"
        params = {
            'ItemM_page': page,
            'ajax': 'item-m-grid',
            'r': 'transaction/itemSetting/index'
        }
        
        try:
            response = session.get(url, params=params, timeout=10)
            parser = ProjectTableParser()
            parser.feed(response.text)
            
            projects = []
            for row in parser.rows:
                if len(row) >= 3:
                    project = {
                        '序号': row[0].strip(),
                        '项目代码': row[1].strip(),
                        '项目名称': row[2].strip(),
                        '分数': row[3].strip() if len(row) > 3 else '0.00'
                    }
                    projects.append(project)
            
            all_projects_by_page[page] = projects
            
            if projects:
                print(f"第 {page:3d} 页：{len(projects)} 条项目")
                print(f"  首条：{projects[0]['项目代码']} - {projects[0]['项目名称'][:30]}")
                print(f"  末条：{projects[-1]['项目代码']} - {projects[-1]['项目名称'][:30]}")
            else:
                print(f"第 {page:3d} 页：❌ 无数据")
            
            print()
        except Exception as e:
            print(f"第 {page:3d} 页：❌ 错误 - {e}\n")
    
    # 检查是否有重复数据
    print("\n" + "="*70)
    print("🔍 检查数据是否重复")
    print("="*70 + "\n")
    
    first_page_codes = set(p['项目代码'] for p in all_projects_by_page.get(1, []))
    last_page_codes = set(p['项目代码'] for p in all_projects_by_page.get(242, []))
    
    print(f"第 1 页的项目代码：{first_page_codes}")
    print(f"第 242 页的项目代码：{last_page_codes}")
    
    if first_page_codes == last_page_codes:
        print("\n❌ 数据完全重复！每页都是同样的 10 条项目")
    else:
        print("\n✅ 数据不同！AJAX 分页有效")
    
    session.close()

if __name__ == "__main__":
    test_ajax_pagination()
