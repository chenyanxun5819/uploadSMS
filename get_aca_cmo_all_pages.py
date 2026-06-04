#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
获取 ACA CMO 项目的全部资料（所有20页）
倒序排列显示
"""

import requests
import time
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


class PaginationParser(HTMLParser):
    """分页链接解析器"""
    def __init__(self):
        super().__init__()
        self.pagination_links = []
        self.in_pagination = False
    
    def handle_starttag(self, tag, attrs):
        if tag == "ul":
            # 检查是否是分页导航
            attrs_dict = dict(attrs)
            if "pagination" in attrs_dict.get("class", ""):
                self.in_pagination = True
        elif tag == "a" and self.in_pagination:
            attrs_dict = dict(attrs)
            href = attrs_dict.get("href", "")
            if href:
                self.pagination_links.append(href)
    
    def handle_endtag(self, tag):
        if tag == "ul":
            self.in_pagination = False


def get_all_aca_cmo_projects():
    """获取 ACA CMO 项目的全部资料（所有20页）"""
    
    print("=" * 80)
    print("🔍 获取 ACA CMO 项目的全部资料（所有 20 页）")
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
    
    # 2. 访问项目设置页面并搜索 ACA CMO
    print("📍 第二步: 搜索 ACA CMO 项目")
    print("-" * 80)
    
    all_projects = []
    page_count = 0
    current_url = ITEM_SETTING_PAGE
    
    try:
        while current_url and page_count < 30:  # 最多获取30页，防止无限循环
            page_count += 1
            print(f"⏳ 正在获取第 {page_count} 页...")
            
            try:
                response = session.get(current_url, timeout=10)
                if response.status_code != 200:
                    print(f"❌ 页面加载失败: {response.status_code}")
                    break
            except Exception as e:
                print(f"❌ 页面加载异常: {e}")
                break
            
            # 首页需要进行搜索
            if page_count == 1:
                print(f"   → 搜索关键字: ACA CMO")
                # 在搜索框中输入 ACA CMO
                search_data = {
                    'ItemSetting[item_code]': 'ACA CMO',
                }
                # 尝试用 POST 搜索
                try:
                    response = session.post(ITEM_SETTING_PAGE, data=search_data, timeout=10)
                except:
                    pass
                
                # 重新获取搜索结果页面
                response = session.get(ITEM_SETTING_PAGE, timeout=10)
            
            # 解析表格数据
            parser = TableParser()
            parser.feed(response.text)
            rows = parser.rows
            
            # 从第3行开始提取项目数据（跳过表头和搜索框）
            page_projects = 0
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
                    except:
                        continue
            
            print(f"   ✓ 本页找到 {page_projects} 个 ACA CMO 项目")
            
            # 查找下一页链接
            next_page_found = False
            
            # 方法1: 查找分页导航中的"下一页"链接
            import re
            
            # 查找所有分页链接
            # 通常分页链接格式为: ?r=transaction/itemSetting/index&page=2
            page_links = re.findall(r'page[=&]\d+', response.text)
            
            if page_links:
                # 获取当前页码
                current_page_match = re.search(r'page[=&](\d+)', current_url)
                if current_page_match:
                    current_page = int(current_page_match.group(1))
                else:
                    current_page = 1
                
                # 构造下一页的 URL
                next_page = current_page + 1
                if f'page={next_page}' in response.text or f'page&{next_page}' in response.text:
                    # 查找完整的下一页链接
                    next_page_match = re.search(
                        rf'(?:href=["\'])?([^"\'>\s]*\?[^"\'>\s]*page[=&]{next_page}[^"\'>\s]*)',
                        response.text
                    )
                    if next_page_match:
                        next_url = next_page_match.group(1)
                        if not next_url.startswith('http'):
                            next_url = 'http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index&page=' + str(next_page)
                        current_url = next_url
                        next_page_found = True
                        print()
                        time.sleep(0.5)
                    else:
                        # 尝试直接构造下一页 URL
                        current_url = ITEM_SETTING_PAGE + f'&page={next_page}&ItemSetting%5Bitem_code%5D=ACA%20CMO'
                        next_page_found = True
                        print()
                        time.sleep(0.5)
            
            if not next_page_found:
                # 尝试查找分页导航中的下一页按钮
                if 'next' in response.text.lower() or 'page' in response.text.lower():
                    # 尝试查找下一页链接的另一种方式
                    import urllib.parse
                    # 构造下一页搜索 URL
                    next_page_num = page_count + 1
                    current_url = ITEM_SETTING_PAGE + f'&page={next_page_num}'
                    next_page_found = True
                    print()
                    time.sleep(0.5)
                else:
                    break
            
            if not next_page_found:
                break
        
        # 3. 显示结果
        print()
        print("=" * 80)
        print(f"📊 共找到 {len(all_projects)} 个 ACA CMO 项目（来自 {page_count} 页）")
        print("=" * 80)
        print()
        
        if len(all_projects) == 0:
            print("⚠️  未找到任何 ACA CMO 项目")
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
        
    except Exception as e:
        print(f"❌ 处理异常: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    get_all_aca_cmo_projects()
