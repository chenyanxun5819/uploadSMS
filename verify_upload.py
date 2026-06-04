#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证上传的成绩数据
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
    session.post(LOGIN_URL, data=login_data, timeout=15, allow_redirects=True)
    
    print("=" * 100)
    print("验证上传的成绩数据")
    print("=" * 100)
    
    # 导航到学生成绩列表页面
    list_url = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/index"
    
    resp = session.get(list_url, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 查找最近的记录
    # 通常会在表格中显示最新的成绩记录
    rows = soup.select('table tbody tr')
    
    print(f"\n找到 {len(rows)} 条记录\n")
    
    # 显示前几条最新的记录
    target_students = {
        '24177': 'J3B 林皓宇',
        '23121': 'S1B 林芊悦',
        '23073': 'S1B 卢旖',
    }
    
    print("查找上传的学生记录:")
    print("=" * 100)
    
    found_count = 0
    
    for row in rows[:50]:  # 查看前 50 条记录
        cells = row.select('td')
        if len(cells) > 0:
            row_text = ' '.join([cell.get_text(strip=True) for cell in cells[:6]])
            
            # 检查是否是我们要找的学生
            for student_id, student_name in target_students.items():
                if student_id in row_text or student_name.split()[1] in row_text:
                    print(f"\n✓ 找到: {student_name}")
                    print(f"  记录: {row_text[:100]}")
                    found_count += 1
    
    print("\n" + "=" * 100)
    print(f"验证结果: {found_count}/3 学生成功上传")
    print("=" * 100)

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
