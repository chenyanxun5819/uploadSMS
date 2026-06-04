#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
详细调试：查看 SMS 系统中的所有学生数据
"""

import sys
from pathlib import Path
import requests
from bs4 import BeautifulSoup

# 禁用 SSL 警告
requests.packages.urllib3.disable_warnings()

print("=" * 100)
print("详细调试：SMS 系统学生数据")
print("=" * 100)

# 登录信息
username = "schhs334"
password = "schhs334"

try:
    session = requests.Session()
    session.verify = False
    
    print("\n[1] 正在登录 SMS 系统...")
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
    
    # 获取活动页面
    print("\n[2] 正在获取 SMS 系统中的所有学生数据...")
    ACTIVITY_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"
    resp = session.get(ACTIVITY_PAGE, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 提取所有学生
    all_students = []
    for link in soup.select('a[data-student_id]'):
        student = {
            'internal_id': link.get('data-student_id'),
            'student_no': link.get('data-student_no'),
            'name': link.get('data-student_name'),
            'class_id': link.get('data-class_id'),
            'class_name': link.get('data-class_name'),
        }
        all_students.append(student)
    
    print(f"  ✓ 共获取 {len(all_students)} 条学生记录")
    
    # 查找目标学生
    print("\n[3] 查找目标学生...")
    target_students = [
        ('J3B', '24177'),
        ('S1B', '23121'),
        ('S1B', '23073'),
    ]
    
    for target_class, target_student_no in target_students:
        print(f"\n  查找: {target_class} - {target_student_no}")
        
        found = False
        for student in all_students:
            if student['student_no'] == target_student_no:
                print(f"    ✓ 找到!")
                print(f"      internal_id: {student['internal_id']}")
                print(f"      name: {student['name']}")
                print(f"      class_id: {student['class_id']}")
                print(f"      class_name: {student['class_name']}")
                found = True
                break
        
        if not found:
            print(f"    ✗ 未找到")
            print(f"      搜索 SMS 系统中是否有类似的学号...")
            
            # 搜索类似的学号
            similar = [s for s in all_students if target_student_no in s['student_no']]
            if similar:
                print(f"      找到 {len(similar)} 条相似的学号:")
                for s in similar[:3]:
                    print(f"        - {s['student_no']} ({s['class_name']}) - internal_id: {s['internal_id']}")
            
            # 搜索该班级的学生
            print(f"      搜索 {target_class} 班的学生...")
            class_students = [s for s in all_students if target_class in s.get('class_name', '')]
            if class_students:
                print(f"      {target_class} 班共有 {len(class_students)} 个学生")
                for s in class_students[:5]:
                    print(f"        - {s['student_no']} - {s['name']} (internal_id: {s['internal_id']})")
                if len(class_students) > 5:
                    print(f"        ... 还有 {len(class_students) - 5} 个")
            else:
                print(f"      ✗ 在 SMS 系统中未找到 {target_class} 班的任何学生")
    
    print("\n" + "=" * 100)
    print("分析:")
    print("=" * 100)
    print("""
可能的原因：
1. Excel 中的学号与 SMS 系统中的学号不一致
2. 学生的班级名称不同（例如 S1B 在 SMS 中可能叫 S1B 或其他名称）
3. SMS 系统中某些班级的学生数据还未更新

建议：
1. 手动在 SMS 网页上验证这些学生是否存在
2. 检查 Excel 中的学号是否正确
3. 如果学号不对，需要更新 Excel 中的数据
""")
    
except Exception as e:
    print(f"  ✗ 错误: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("=" * 100)
