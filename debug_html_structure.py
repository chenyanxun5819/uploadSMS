#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
改进的学生数据提取 - 正确解析 HTML
"""

import sys
import requests
from bs4 import BeautifulSoup

requests.packages.urllib3.disable_warnings()

print("=" * 100)
print("改进的学生数据提取 - 正确解析 HTML")
print("=" * 100)

username = "schhs334"
password = "schhs334"

try:
    session = requests.Session()
    session.verify = False
    
    # 1. 登录
    print("\n[1] 正在登录...")
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    login_data = {
        'LoginForm[username]': username,
        'LoginForm[password]': password,
        'login-button': 'login'
    }
    
    session.get(LOGIN_URL, timeout=15)
    resp = session.post(LOGIN_URL, data=login_data, timeout=15, allow_redirects=True)
    
    if 'login' in resp.url.lower():
        print("  ✗ 登录失败")
        sys.exit(1)
    
    print("  ✓ 登录成功")
    
    # 2. 获取班级的学生数据
    print("\n[2] 正在获取班级学生数据...")
    
    item_id = "2444"
    date_str = "2026-02-06"
    
    # 重点班级：734 (S1B), 706 (?), 等等
    target_class_ids = ['734', '706']  # S1B 和 J3B 应该在其中
    
    all_students_by_class = {}
    
    for class_id in target_class_ids:
        print(f"\n  班级 ID: {class_id}")
        
        ajax_url = "http://sms.chhsban.edu.my/sms/index.php"
        ajax_params = {
            'r': 'transaction/studentPerformance/update',
            'StudentPerformanceM[class_id]': class_id,
            'StudentPerformanceM[item_id]': item_id,
            'ajax': 'student-grid',
            'date': date_str,
            'item_id': item_id,
        }
        
        try:
            resp = session.get(ajax_url, params=ajax_params, timeout=10)
            
            if resp.status_code == 200:
                # 保存 HTML 用于调试
                with open(f'class_{class_id}_response.html', 'w', encoding='utf-8') as f:
                    f.write(resp.text)
                print(f"    ✓ 已保存 HTML 到 class_{class_id}_response.html")
                
                soup = BeautifulSoup(resp.text, 'html.parser')
                
                # 查找所有可能包含学生信息的元素
                print(f"    搜索学生数据...")
                
                # 方法 1: 查找 <tr> 行，可能每行是一个学生
                rows = soup.select('table tbody tr')
                print(f"    找到 {len(rows)} 行")
                
                if rows:
                    students = []
                    for row in rows[:5]:  # 先看前 5 行
                        # 从行中提取数据
                        # 可能的格式: <td> 包含学号、名字等
                        tds = row.select('td')
                        if tds:
                            row_text = [td.get_text(strip=True) for td in tds]
                            print(f"      行数据: {row_text}")
                            
                            # 尝试从 data-* 属性提取
                            data_attrs = row.attrs
                            if 'class' in data_attrs and 'data-student_id' in row.attrs:
                                student = {
                                    'internal_id': row.get('data-student_id'),
                                    'student_no': row.get('data-student_no'),
                                    'name': row.get('data-student_name'),
                                    'class_id': class_id,
                                }
                                print(f"      数据属性: {student}")
                                students.append(student)
                    
                    all_students_by_class[class_id] = students
                    
                    # 也尝试查找 <a> 标签
                    links = soup.select('a[data-student_id]')
                    if links:
                        print(f"    找到 {len(links)} 个 <a> 标签 (data-student_id)")
                        for link in links[:3]:
                            print(f"      - {link.get('data-student_no')}: {link.get('data-student_name')}")
                
        except Exception as e:
            print(f"    ✗ 错误: {e}")
    
    print("\n" + "=" * 100)
    print("调试完成 - 请检查 class_*_response.html 文件中的 HTML 结构")
    print("=" * 100)

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
