#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正确的学生提取方法：
1. 获取所有班级列表
2. 对每个班级发送 AJAX 请求获取学生
3. 合并所有学生数据
"""

import sys
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlencode

requests.packages.urllib3.disable_warnings()

print("=" * 100)
print("正确的学生数据提取 - 按班级逐一获取")
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
    
    # 2. 获取活动页面，提取班级列表
    print("\n[2] 正在获取班级列表...")
    ACTIVITY_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"
    resp = session.get(ACTIVITY_PAGE, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 尝试找到班级选择的 select 元素
    # XPath: /html/body/div[2]/div[2]/div[2]/div[2]/div[2]/form/div[7]/div/div[1]/table/tbody/tr[2]/th[2]/select
    # 这通常是动态加载的，我们先试试从项目页面中找班级选项
    
    print("  搜索班级选择下拉菜单...")
    
    # 尝试找所有 select 元素
    selects = soup.select('select')
    print(f"    找到 {len(selects)} 个 select 元素")
    
    for idx, sel in enumerate(selects):
        options = sel.select('option')
        if len(options) > 5 and len(options) < 50:  # 班级通常 10-30 个
            print(f"    [可能是班级选择] select #{idx}: {len(options)} 个选项")
            for opt in options[:5]:
                print(f"      - {opt.get('value')}: {opt.get_text(strip=True)}")
            if len(options) > 5:
                print(f"      ... 还有 {len(options) - 5} 个")
    
    # 3. 尝试通过 AJAX 获取班级的学生
    print("\n[3] 尝试通过 AJAX 获取学生数据...")
    print("  (这需要知道班级 ID，我们先尝试常见的班级 ID)")
    
    # 从之前的结果推断，S1B 班级 ID 可能是 734（之前找到的），让我们试试
    # J3B 应该有一个不同的 class_id
    
    # 先尝试用项目信息来获取学生列表
    # 根据 /transaction/studentPerformance/update 的 AJAX 参数：
    # StudentPerformanceM[class_id]=701&StudentPerformanceM[item_id]=2444&ajax=student-grid
    
    item_id = "2444"  # 从之前的 curl 中得到
    date_str = "2026-02-06"
    
    # 尝试几个常见的班级 ID
    test_class_ids = ['734', '700', '701', '702', '703', '704', '705', '706']
    
    all_students = {}
    
    for class_id in test_class_ids:
        print(f"\n  尝试班级 ID: {class_id}")
        
        # 构建 AJAX 请求 URL
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
            
            if resp.status_code == 200 and resp.text.strip():
                soup_ajax = BeautifulSoup(resp.text, 'html.parser')
                
                # 从 HTML 中提取学生数据
                # 查找包含 data-student_id 的元素
                students = soup_ajax.select('[data-student_id]')
                
                if students:
                    print(f"    ✓ 找到 {len(students)} 个学生")
                    all_students[class_id] = []
                    
                    for student_elem in students[:3]:
                        student_no = student_elem.get('data-student_no')
                        internal_id = student_elem.get('data-student_id')
                        name = student_elem.get('data-student_name')
                        class_name = student_elem.get('data-class_name')
                        
                        all_students[class_id].append({
                            'student_no': student_no,
                            'internal_id': internal_id,
                            'name': name,
                            'class_name': class_name,
                        })
                        print(f"      - {student_no}: {name} ({class_name})")
                    
                    if len(students) > 3:
                        print(f"      ... 还有 {len(students) - 3} 个")
                else:
                    print(f"    ✗ 未找到学生")
        
        except Exception as e:
            print(f"    ✗ 错误: {e}")
    
    print("\n" + "=" * 100)
    print("总结:")
    print("=" * 100)
    
    total_students = sum(len(v) for v in all_students.values())
    print(f"\n找到 {len(all_students)} 个班级，共 {total_students} 个学生\n")
    
    for class_id, students in all_students.items():
        if students:
            print(f"班级 ID {class_id}:")
            for s in students:
                print(f"  - {s['student_no']}: {s['name']}")

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 100)
