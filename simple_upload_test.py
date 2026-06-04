#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
纯 requests 版本的上传测试
"""

import requests
from bs4 import BeautifulSoup

requests.packages.urllib3.disable_warnings()

username = "schhs334"
password = "schhs334"

try:
    session = requests.Session()
    session.verify = False
    
    print("=" * 100)
    print("纯 requests 版本上传测试")
    print("=" * 100)
    
    # 1. 登录
    print("\n1. 登录...")
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    login_data = {
        'LoginForm[username]': username,
        'LoginForm[password]': password,
        'login-button': 'login'
    }
    
    session.get(LOGIN_URL, timeout=15)
    resp = session.post(LOGIN_URL, data=login_data, timeout=15, allow_redirects=True)
    print(f"   登录响应: {resp.status_code}")
    
    # 2. 构建上传表单数据
    print("\n2. 构建表单数据...")
    
    post_data = {
        'StudentPerformanceM[year]': '2026',
        'StudentPerformanceM[semester]': '1',
        'StudentPerformanceM[date]': '2026-02-06',
        'StudentPerformanceM[item_id]': '2444',  # ACA CMO207
        'filterS': 'class',
        'class_id': '726',  # J3B
        'club_id': '53',
        'StudentM[student_no]': '',
        'StudentM[name]': '',
        'StudentM[cname]': '',
        'StudentM[class_name]': '',
        'yt1': '',
    }
    
    # 添加学生数据
    # 学生 1: 24177 (internal_id: 7816)
    post_data['StudentPerformanceM[inputperformance][7816][class_id]'] = '726'
    post_data['StudentPerformanceM[inputperformance][7816][type_of_bonus]'] = '1'
    post_data['StudentPerformanceM[inputperformance][7816][mark]'] = '0.00'
    post_data['StudentPerformanceM[inputperformance][7816][remark]'] = ''
    
    # 学生 2: 23121 (internal_id: 7244)
    post_data['StudentPerformanceM[inputperformance][7244][class_id]'] = '734'
    post_data['StudentPerformanceM[inputperformance][7244][type_of_bonus]'] = '1'
    post_data['StudentPerformanceM[inputperformance][7244][mark]'] = '0.00'
    post_data['StudentPerformanceM[inputperformance][7244][remark]'] = ''
    
    # 学生 3: 23073 (internal_id: 6729)
    post_data['StudentPerformanceM[inputperformance][6729][class_id]'] = '734'
    post_data['StudentPerformanceM[inputperformance][6729][type_of_bonus]'] = '1'
    post_data['StudentPerformanceM[inputperformance][6729][mark]'] = '0.00'
    post_data['StudentPerformanceM[inputperformance][6729][remark]'] = ''
    
    print(f"   表单字段数: {len(post_data)}")
    
    # 3. 提交表单
    print("\n3. 提交表单...")
    ACTIVITY_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"
    
    resp = session.post(ACTIVITY_PAGE, data=post_data, timeout=15, allow_redirects=True)
    
    print(f"   POST 响应: {resp.status_code}")
    
    if resp.status_code == 200:
        # 检查响应中是否有成功信息
        if 'success' in resp.text.lower() or '成功' in resp.text:
            print("   ✓ 表单已提交")
        else:
            print("   提交状态: 未知")
        
        # 4. 等一下，然后查询最新的记录
        print("\n4. 查询最新的记录...")
        
        import time
        time.sleep(2)  # 等待 2 秒让数据保存
        
        # 访问结果页面或列表页面
        resp = session.get(ACTIVITY_PAGE, timeout=15)
        soup = BeautifulSoup(resp.text, 'html.parser')
        
        # 查找是否有我们输入的学生
        print("   查找上传的学生数据...")
        
        target_students = ['24177', '23121', '23073']
        found = 0
        
        for student_id in target_students:
            if student_id in resp.text:
                # 再检查是否在表格中
                links = soup.select('a[data-student_no]')
                for link in links:
                    if link.get('data-student_no') == student_id:
                        # 检查是否有标记
                        parent = link.find_parent('tr')
                        if parent:
                            row_text = parent.get_text()
                            # 如果行中有成绩或标记，说明已保存
                            print(f"   ✓ 学生 {student_id} 已处理")
                            found += 1
                            break
        
        print(f"\n   找到 {found} 个学生")
    else:
        print(f"   ✗ POST 失败: {resp.status_code}")
    
    print("\n" + "=" * 100)
    print("测试完成")
    print("=" * 100)

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
