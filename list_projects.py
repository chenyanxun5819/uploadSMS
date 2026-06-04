#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文件 - 列出系统中所有的项目名称
使用账号/密码: schhs334/schhs334
不需要额外依赖库
"""

import requests
import re
import time
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from html.parser import HTMLParser

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


def get_all_projects():
    """获取所有项目列表"""
    
    print("=" * 70)
    print("📋 获取系统中所有项目名称")
    print("=" * 70)
    print()
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    ITEM_SETTING_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"
    
    # 创建会话
    session = requests.Session()
    session.verify = False
    
    # 1. 登入系统
    print("📍 登入系统...")
    print("-" * 70)
    
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
    
    # 2. 获取项目列表页面
    print("📍 获取项目列表...")
    print("-" * 70)
    
    try:
        response = session.get(ITEM_SETTING_PAGE, timeout=10)
        if response.status_code == 200:
            print("✅ 页面加载成功")
        else:
            print(f"❌ 无法加载页面: {response.status_code}")
            return
    except Exception as e:
        print(f"❌ 页面加载失败: {e}")
        return
    
    print()
    time.sleep(1)
    
    # 3. 解析HTML获取项目信息
    print("📍 解析项目数据...")
    print("-" * 70)
    
    try:
        # 使用内置的 HTMLParser
        parser = TableParser()
        parser.feed(response.text)
        
        rows = parser.rows
        print(f"✅ 找到表格，共 {len(rows)} 行")
        print()
        
        # 提取项目数据（跳过表头行）
        projects = []
        
        # 第一行通常是表头，第二行是搜索框，从第三行开始是数据
        for row in rows[2:]:
            if len(row) >= 3:
                try:
                    project_code = row[1].strip()
                    project_name = row[2].strip()
                    
                    if project_code:  # 只记录有项目代码的行
                        projects.append({
                            'code': project_code,
                            'name': project_name
                        })
                except:
                    continue
        
        # 4. 显示结果
        print()
        print("=" * 70)
        print(f"📊 共找到 {len(projects)} 个项目")
        print("=" * 70)
        print()
        
        if len(projects) == 0:
            print("⚠️  未找到任何项目")
            return
        
        # 倒序排列（最新的在前）
        projects_reversed = list(reversed(projects))
        
        print("项目名称列表 (从最新开始倒序排列):")
        print("-" * 70)
        
        for i, proj in enumerate(projects_reversed, 1):
            print(f"{i:3}. {proj['name']}")
        
        print()
        print("=" * 70)
        print(f"✅ 共列出 {len(projects)} 个项目名称")
        print("=" * 70)
        
        # 保存到文件
        with open('项目列表.txt', 'w', encoding='utf-8') as f:
            f.write("系统中的所有项目名称 (从最新开始倒序排列)\n")
            f.write("=" * 70 + "\n\n")
            for i, proj in enumerate(projects_reversed, 1):
                f.write(f"{i:3}. {proj['name']}\n")
        
        print()
        print("💾 项目列表已保存到: 项目列表.txt")
        
    except Exception as e:
        print(f"❌ 解析失败: {e}")
        import traceback
        traceback.print_exc()
        return


if __name__ == "__main__":
    get_all_projects()
