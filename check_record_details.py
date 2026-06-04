#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
查看最新一条 ACA CMO207 记录的详情
"""

import sys
import requests
from bs4 import BeautifulSoup
import re

requests.packages.urllib3.disable_warnings()

username = "schhs334"
password = "schhs334"

try:
    session = requests.Session()
    session.verify = False
    
    # 登录
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    login_data = {
        'LoginForm[username]': username,
        'LoginForm[password]': password,
        'login-button': 'login'
    }
    
    session.get(LOGIN_URL, timeout=15)
    session.post(LOGIN_URL, data=login_data, timeout=15, allow_redirects=True)
    
    print("=" * 100)
    print("查看最新的 ACA CMO207 记录详情")
    print("=" * 100)
    
    # 获取记录列表
    list_url = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/index"
    
    resp = session.get(list_url, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 查找所有记录行
    rows = soup.select('table tbody tr')
    
    # 查找 ACA CMO 的记录
    aca_cmo_rows = []
    for idx, row in enumerate(rows):
        text = row.get_text()
        if 'ACA' in text and 'CMO' in text:
            aca_cmo_rows.append((idx, row))
    
    print(f"\n找到 {len(aca_cmo_rows)} 条 ACA CMO 记录\n")
    
    # 查看每条记录的链接
    for idx, row in aca_cmo_rows[:3]:  # 看最新的 3 条
        cells = row.select('td')
        
        # 查找可点击的链接
        link = row.select_one('a[href*="view"]')
        
        if link:
            href = link.get('href', '')
            text = ' | '.join([cell.get_text(strip=True)[:35] for cell in cells[:6]])
            
            print(f"记录 {idx}: {text}")
            
            # 如果有 view 链接，就点进去看详情
            if 'view' in href and 'id=' in href:
                # 提取 ID
                id_match = re.search(r'id=(\d+)', href)
                if id_match:
                    record_id = id_match.group(1)
                    
                    # 获取详情页面
                    detail_url = f"http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/view&id={record_id}"
                    
                    detail_resp = session.get(detail_url, timeout=15)
                    detail_soup = BeautifulSoup(detail_resp.text, 'html.parser')
                    
                    # 查找学生数据表格
                    student_tables = detail_soup.select('table')
                    
                    target_students = ['24177', '23121', '23073']
                    
                    print(f"\n  详情 (ID: {record_id}):")
                    
                    # 在详情页面中查找学生
                    found_students = []
                    for student_id in target_students:
                        if student_id in detail_resp.text:
                            found_students.append(student_id)
                            print(f"    ✓ 找到学生: {student_id}")
                    
                    if not found_students:
                        print(f"    - 未找到任何目标学生")
                    
                    print()
    
    # 如果没有找到链接，尝试获取最新上传的记录
    print("\n查询最新上传日期是 2026-02-06 的记录...")
    
    # 直接通过 URL 查询特定日期的记录
    search_url = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/index&StudentPerformanceSearch[date]=2026-02-06"
    
    resp = session.get(search_url, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    rows = soup.select('table tbody tr')
    print(f"找到 {len(rows)} 条 2026-02-06 的记录\n")
    
    for row in rows[:5]:
        text = ' | '.join([cell.get_text(strip=True)[:40] for cell in row.select('td')[:6]])
        print(f"  {text}")

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
