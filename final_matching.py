#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
最终的学生匹配 - 完整版
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
    print("最终学生匹配")
    print("=" * 100)
    
    item_id = "2444"
    date_str = "2026-02-06"
    
    # 已知的班级 ID 映射
    class_ids = {
        'J3B': '726',
        'S1B': '734',
    }
    
    # 从 Excel 中需要找的学生
    target_students = [
        {'class': 'J3B', 'student_id': '24177', 'name': '林皓宇'},
        {'class': 'S1B', 'student_id': '23121', 'name': '林芊悦'},
        {'class': 'S1B', 'student_id': '23073', 'name': '卢旖'},
    ]
    
    print("\n需要匹配的学生:")
    for s in target_students:
        print(f"  {s['class']:5} - {s['student_id']:10} - {s['name']}")
    
    # 获取每个班级的所有学生
    print("\n" + "=" * 100)
    print("从 SMS 系统获取学生数据")
    print("=" * 100)
    
    all_students = {}
    
    for class_name, class_id in class_ids.items():
        print(f"\n获取 {class_name} (ID: {class_id})...")
        
        ajax_url = "http://sms.chhsban.edu.my/sms/index.php"
        ajax_params = {
            'r': 'transaction/studentPerformance/update',
            'StudentPerformanceM[class_id]': class_id,
            'StudentPerformanceM[item_id]': item_id,
            'ajax': 'student-grid',
            'date': date_str,
            'item_id': item_id,
        }
        
        resp = session.get(ajax_url, params=ajax_params, timeout=10)
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        # 获取所有学生
        links = soup.select('a[data-student_id]')
        
        for link in links:
            student_no = link.get('data-student_no')
            internal_id = link.get('data-student_id')
            name = link.get('data-student_name')
            
            all_students[student_no] = {
                'internal_id': internal_id,
                'name': name,
                'class_name': class_name,
                'class_id': class_id,
            }
        
        print(f"  ✓ 获取 {len(links)} 个学生")
    
    # 匹配学生
    print("\n" + "=" * 100)
    print("匹配结果")
    print("=" * 100)
    
    matched_count = 0
    unmatched = []
    
    for target in target_students:
        student_id = target['student_id']
        class_name = target['class']
        name = target['name']
        
        print(f"\n查找: {class_name:5} - {student_id:10} - {name}")
        
        if student_id in all_students:
            sms_student = all_students[student_id]
            print(f"  ✓ 找到!")
            print(f"    internal_id: {sms_student['internal_id']}")
            print(f"    SMS 名字: {sms_student['name']}")
            print(f"    SMS 班级: {sms_student['class_name']}")
            matched_count += 1
        else:
            print(f"  ✗ 未找到")
            unmatched.append(f"{class_name} {student_id}")
            
            # 搜索类似的学号
            similar = [sno for sno in all_students.keys() if student_id in sno]
            if similar:
                print(f"    类似的学号: {similar[:3]}")
            
            # 搜索该班级中包含该名字的学生
            for sno, student in all_students.items():
                if student['class_name'] == class_name and name in student['name']:
                    print(f"    找到名字相似的学生: {sno} - {student['name']}")
    
    print("\n" + "=" * 100)
    print(f"最终结果: {matched_count}/3 匹配成功")
    print("=" * 100)
    
    if unmatched:
        print("\n未匹配的学生:")
        for s in unmatched:
            print(f"  - {s}")

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
