#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
找到 J3B 班级的 ID
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
    print("扫描所有班级找 J3B")
    print("=" * 100)
    
    item_id = "2444"
    date_str = "2026-02-06"
    
    # 扫描班级 ID 700-750
    found_classes = {}
    
    for class_id in range(700, 751):
        class_id_str = str(class_id)
        
        ajax_url = "http://sms.chhsban.edu.my/sms/index.php"
        ajax_params = {
            'r': 'transaction/studentPerformance/update',
            'StudentPerformanceM[class_id]': class_id_str,
            'StudentPerformanceM[item_id]': item_id,
            'ajax': 'student-grid',
            'date': date_str,
            'item_id': item_id,
        }
        
        try:
            resp = session.get(ajax_url, params=ajax_params, timeout=10)
            
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'html.parser')
                
                # 查找学生数据
                links = soup.select('a[data-student_id]')
                
                if links and len(links) > 0:
                    # 获取班级名称（从第一个学生的数据获取）
                    first_student = links[0]
                    class_name = first_student.get('data-class_name', 'Unknown')
                    
                    found_classes[class_id_str] = {
                        'class_name': class_name,
                        'student_count': len(links),
                        'first_students': [l.get('data-student_no') for l in links[:3]]
                    }
                    
                    print(f"班级 ID {class_id_str:3}: {class_name:10} ({len(links):2} 学生) - 样本: {found_classes[class_id_str]['first_students']}")
                    
                    if 'J3B' in class_name:
                        print(f"\n✓✓✓ 找到 J3B！班级 ID = {class_id_str}")
        
        except:
            pass
    
    print("\n" + "=" * 100)
    print("找到的所有班级:")
    print("=" * 100)
    
    for class_id, info in sorted(found_classes.items()):
        print(f"班级 ID {class_id}: {info['class_name']:10} ({info['student_count']:2} 学生)")

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
