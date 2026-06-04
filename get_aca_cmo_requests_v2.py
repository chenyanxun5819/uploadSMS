#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
获取 ACA CMO 项目的全部资料（所有页）- 使用 requests 库
通过分析 URL 参数来获取所有分页数据
"""

import requests
import time
import re
from html.parser import HTMLParser
from requests.packages.urllib3.exceptions import InsecureRequestWarning

# 禁用 SSL 警告
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


class TableParser(HTMLParser):
    """简单的表格解析器"""
    def __init__(self):
        super().__init__()
        self.rows = []
        self.current_row = []
        self.in_table = False
        self.in_td = False
        self.current_cell = ""
    
    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.in_table = True
        elif tag == "tr" and self.in_table:
            self.current_row = []
        elif tag in ["td", "th"] and self.in_table:
            self.in_td = True
            self.current_cell = ""
    
    def handle_endtag(self, tag):
        if tag == "table":
            self.in_table = False
        elif tag == "tr" and self.in_table:
            if self.current_row:
                self.rows.append(self.current_row)
        elif tag in ["td", "th"] and self.in_table:
            self.in_td = False
            self.current_row.append(self.current_cell.strip())
    
    def handle_data(self, data):
        if self.in_td:
            self.current_cell += data


def get_aca_cmo_all_pages_requests():
    """使用 requests 库获取 ACA CMO 项目的全部资料"""
    
    print("=" * 80)
    print("🔍 获取 ACA CMO 项目的全部资料（Requests 库版本）")
    print("=" * 80)
    print()
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    ITEM_SETTING_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"
    
    # 创建会话
    session = requests.Session()
    session.verify = False
    
    # 1. 登入系统
    print("📍 第一步: 登入系统")
    print("-" * 80)
    
    login_data = {
        'LoginForm[username]': 'schhs334',
        'LoginForm[password]': 'schhs334',
        'login-button': 'login'
    }
    
    try:
        response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
        
        if 'login' not in response.url.lower():
            print("✅ 登入成功")
        else:
            print("❌ 登入失败")
            return
    except Exception as e:
        print(f"❌ 登入失败: {e}")
        return
    
    print()
    time.sleep(1)
    
    # 2. 首先获取搜索结果的第一页，确定项目代码的确切参数名
    print("📍 第二步: 搜索 ACA CMO 项目（获取第一页）")
    print("-" * 80)
    
    # 尝试不同的搜索参数格式
    search_params = {
        'ItemSetting[item_code]': 'ACA CMO'
    }
    
    try:
        response = session.get(ITEM_SETTING_PAGE, params=search_params, timeout=10)
        print(f"   → 搜索 URL: {response.url}")
        print("✅ 第一页加载成功")
    except Exception as e:
        print(f"❌ 第一页加载失败: {e}")
        return
    
    print()
    time.sleep(1)
    
    # 3. 收集所有页面的数据
    print("📍 第三步: 遍历所有页面")
    print("-" * 80)
    
    all_projects = []
    page_count = 0
    max_pages = 25
    
    for page in range(1, max_pages + 1):
        page_count = page
        print(f"⏳ 正在获取第 {page} 页...")
        
        try:
            # 构造带分页参数的请求
            search_params_with_page = {
                'ItemSetting[item_code]': 'ACA CMO',
                'page': page
            }
            
            response = session.get(ITEM_SETTING_PAGE, params=search_params_with_page, timeout=10)
            
            if response.status_code != 200:
                print(f"   ❌ 页面加载失败: {response.status_code}")
                break
            
            # 解析表格数据
            parser = TableParser()
            parser.feed(response.text)
            rows = parser.rows
            
            if len(rows) < 3:  # 没有数据行
                print(f"   → 已到达最后一页（无数据）")
                break
            
            # 从第3行开始提取项目数据（跳过表头和搜索框）
            page_projects = 0
            page_has_data = False
            
            for row in rows[2:]:
                if len(row) >= 3:
                    try:
                        project_code = row[1].strip()
                        project_name = row[2].strip()
                        
                        if project_code and 'ACA CMO' in project_code:
                            all_projects.append({
                                'code': project_code,
                                'name': project_name
                            })
                            page_projects += 1
                            page_has_data = True
                    except:
                        continue
            
            if page_has_data:
                print(f"   ✓ 本页找到 {page_projects} 个 ACA CMO 项目")
            else:
                print(f"   → 本页无数据，已到达最后一页")
                break
        
        except Exception as e:
            print(f"   ⚠️  第 {page} 页加载异常: {e}")
            if page == 1:
                break
        
        time.sleep(0.5)
    
    # 4. 显示结果
    print()
    print("=" * 80)
    print(f"📊 共获取 {page_count} 页数据，找到 {len(all_projects)} 个 ACA CMO 项目")
    print("=" * 80)
    print()
    
    if len(all_projects) == 0:
        print("⚠️  未找到任何项目")
        return
    
    # 倒序排列
    all_projects_reversed = list(reversed(all_projects))
    
    print("ACA CMO 项目列表 (倒序排列 - 最新在前):")
    print("-" * 80)
    print()
    
    for i, proj in enumerate(all_projects_reversed, 1):
        print(f"{i:4}. 代码: {proj['code']:<20} | 名称: {proj['name']}")
    
    print()
    print("=" * 80)
    print(f"✅ 共列出 {len(all_projects)} 个 ACA CMO 项目")
    print("=" * 80)
    
    # 保存到文件
    with open('ACA_CMO_全部资料.txt', 'w', encoding='utf-8') as f:
        f.write("ACA CMO 项目全部资料（倒序排列 - 最新在前）\n")
        f.write("=" * 80 + "\n\n")
        for i, proj in enumerate(all_projects_reversed, 1):
            f.write(f"{i:4}. 代码: {proj['code']:<20} | 名称: {proj['name']}\n")
    
    print()
    print("💾 数据已保存到: ACA_CMO_全部资料.txt")


if __name__ == "__main__":
    get_aca_cmo_all_pages_requests()
