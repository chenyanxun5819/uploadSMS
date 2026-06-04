#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高效自动检测最后一页 - 使用二分查找法
"""

import requests
import time
from html.parser import HTMLParser
from requests.packages.urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


class StudentTableParser(HTMLParser):
    """学生成绩表格解析器"""
    def __init__(self):
        super().__init__()
        self.has_data = False
    
    def handle_starttag(self, tag, attrs):
        if tag == "tbody":
            self.has_data = True


def check_page_has_data(session, url, page_num):
    """快速检查某一页是否有数据"""
    try:
        search_params = {
            'StudentPerformance[mark_item]': 'ACA CMO5',
            'page': page_num
        }
        response = session.get(url, params=search_params, timeout=10)
        
        if response.status_code != 200:
            return False
        
        # 快速检查是否有tbody标签（表示有数据）
        parser = StudentTableParser()
        parser.feed(response.text)
        return parser.has_data
    except:
        return False


def find_last_page_binary_search(session, url):
    """
    使用二分查找法自动找出最后一页
    时间复杂度: O(log n) 而不是 O(n)
    """
    print("🔍 自动检测最后一页（二分查找法）...")
    print("-" * 90)
    
    # 第一步: 找上界
    print("⏳ 第1步: 查找数据范围上界...", flush=True)
    upper_bound = 1000
    step = 1000
    
    while check_page_has_data(session, url, upper_bound):
        upper_bound += step
        print(f"  → 第 {upper_bound} 页有数据，继续扩大搜索范围...", flush=True)
        time.sleep(0.2)
        if upper_bound > 100000:  # 防止无限循环
            print("⚠️  数据量过大，超过 100000 页")
            return upper_bound
    
    print(f"  ✓ 上界: {upper_bound} 页\n")
    
    # 第二步: 二分查找最后一页
    print(f"⏳ 第2步: 在 1 ~ {upper_bound} 页间二分查找最后一页...", flush=True)
    
    low = 1
    high = upper_bound
    last_page = 1
    iteration = 0
    
    while low <= high:
        iteration += 1
        mid = (low + high) // 2
        
        print(f"  [{iteration}] 检查第 {mid} 页...", end="", flush=True)
        
        if check_page_has_data(session, url, mid):
            print(" ✓ 有数据", flush=True)
            last_page = mid
            low = mid + 1
        else:
            print(" ✗ 无数据", flush=True)
            high = mid - 1
        
        time.sleep(0.2)
    
    print(f"\n✅ 最后一页是: 第 {last_page} 页\n")
    return last_page


def get_aca_cmo5_student_scores_summary():
    """快速摘要模式 - 只获取最后几页"""
    
    print("=" * 90)
    print("🎓 ACA CMO5 数据自动检测 - 摘要模式（只检测页码，不获取全部数据）")
    print("=" * 90)
    print()
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    PERFORMANCE_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/index"
    
    session = requests.Session()
    session.verify = False
    
    # 1. 登入
    print("📍 第一步: 登入系统")
    print("-" * 90)
    
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
    
    # 2. 自动检测最后一页
    print("📍 第二步: 自动检测最后一页")
    print("-" * 90)
    print()
    
    last_page = find_last_page_binary_search(session, PERFORMANCE_PAGE)
    
    print()
    print("=" * 90)
    print(f"🎯 自动检测结果：ACA CMO5 共有 {last_page} 页数据")
    print("=" * 90)
    print()
    
    # 3. 获取最后一页的数据（作为示例）
    print(f"📍 第三步: 获取最后一页的数据（第 {last_page} 页）")
    print("-" * 90)
    
    try:
        search_params = {
            'StudentPerformance[mark_item]': 'ACA CMO5',
            'page': last_page
        }
        
        response = session.get(PERFORMANCE_PAGE, params=search_params, timeout=10)
        
        # 简单计算行数
        rows_count = response.text.count('<tr>')
        print(f"✓ 第 {last_page} 页包含约 {rows_count - 1} 条记录")
        
    except Exception as e:
        print(f"⚠️  获取最后一页失败: {e}")
    
    print()
    print("=" * 90)
    print("📊 关键信息：")
    print(f"   - 总页数: {last_page} 页")
    print(f"   - 预估总记录数: 约 {last_page * 10} 条（每页约10条）")
    print("=" * 90)
    print()
    print("💡 如需获取所有数据，可运行: python get_aca_cmo5_students.py")


if __name__ == '__main__':
    get_aca_cmo5_student_scores_summary()
