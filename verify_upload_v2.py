#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证上传的成绩数据 - 改进版
直接查询 ACA CMO207 活动的记录
"""

import sys
import requests
from bs4 import BeautifulSoup

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
    resp = session.post(LOGIN_URL, data=login_data, timeout=15, allow_redirects=True)
    
    print("=" * 100)
    print("验证上传的成绩数据 - 查询 ACA CMO207 活动")
    print("=" * 100)
    
    # 直接访问活动详情页面（使用 item_id: 2444）
    # 或者导航到详情页面
    activity_url = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/view&id=something"
    
    # 或者尝试列表页面，搜索 ACA CMO207
    list_url = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/index"
    
    resp = session.get(list_url, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 查找包含 ACA CMO207 或 2444 的记录
    print("\n查找所有记录中 ACA CMO207 的数据...\n")
    
    rows = soup.select('table tbody tr')
    
    target_students = {
        '24177': 'J3B 林皓宇',
        '23121': 'S1B 林芊悦',
        '23073': 'S1B 卢旖',
    }
    
    found_records = []
    
    for row in rows:
        cells = row.select('td')
        if len(cells) >= 4:
            # 获取活动名称
            activity_name = cells[1].get_text(strip=True) if len(cells) > 1 else ""
            
            # 检查是否是 ACA CMO 相关的活动
            if 'CMO' in activity_name or 'ACA' in activity_name or 'JMNC' in activity_name or 'CMO' in cells[0].get_text():
                # 获取完整记录
                record_text = ' '.join([cell.get_text(strip=True)[:50] for cell in cells[:8]])
                
                # 检查是否包含我们的学生
                for student_id, student_name in target_students.items():
                    if student_id in record_text or student_name.split()[1] in record_text:
                        found_records.append({
                            'activity': activity_name,
                            'date': cells[3].get_text(strip=True) if len(cells) > 3 else "",
                            'content': record_text,
                            'student': student_name
                        })
    
    # 也直接查找所有包含学生学号的行
    print("扫描所有行查找学生学号...\n")
    
    found_by_id = {}
    for student_id, student_name in target_students.items():
        for row in rows:
            if student_id in row.get_text():
                cells = row.select('td')
                record_text = ' | '.join([cell.get_text(strip=True)[:40] for cell in cells[:6]])
                found_by_id[student_name] = record_text
                print(f"✓ {student_name}")
                print(f"  {record_text}\n")
    
    if found_by_id:
        print(f"找到 {len(found_by_id)}/3 学生的记录")
    else:
        print("未找到任何学生的具体记录")
        print("\n全部记录列表:")
        for i, row in enumerate(rows[:10]):
            cells = row.select('td')
            if cells:
                text = ' | '.join([cell.get_text(strip=True)[:30] for cell in cells[:5]])
                print(f"  [{i}] {text}")

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
