#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
直接提交表单并检查结果
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'sms_app'))

from core.sms_handler import SMSHandler

handler = SMSHandler()

# 登录
print("=" * 100)
print("测试直接访问上传表单")
print("=" * 100)

try:
    # 手动登录
    print("\n1. 登录...")
    handler.login('schhs334', 'schhs334')
    print("  ✓ 登录成功")
    
    # 获取活动页面
    print("\n2. 获取活动页面...")
    import requests
    from bs4 import BeautifulSoup
    
    resp = handler.session.get(handler.ACTIVITY_PAGE, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 查找是否有表单
    form = soup.select_one('form')
    print(f"  表单: {form is not None}")
    
    # 查找 item_id 选项
    select_element = soup.select_one('select#StudentPerformanceM_item_id')
    if select_element:
        options = select_element.select('option')
        print(f"  活动选项: {len(options)} 个")
        
        # 查找 ACA CMO207
        for option in options:
            text = option.get_text(strip=True)
            value = option.get('value')
            if 'ACA' in text and 'CMO207' in text:
                print(f"  找到: {text} (value: {value})")
    
    # 现在尝试获取学生表格
    print("\n3. 获取学生表格...")
    
    # 发送 AJAX 请求来获取班级 726 (J3B) 的学生
    ajax_url = "http://sms.chhsban.edu.my/sms/index.php"
    ajax_params = {
        'r': 'transaction/studentPerformance/update',
        'StudentPerformanceM[class_id]': '726',
        'StudentPerformanceM[item_id]': '2444',
        'ajax': 'student-grid',
        'date': '2026-02-06',
        'item_id': '2444',
    }
    
    resp = handler.session.get(ajax_url, params=ajax_params, timeout=15)
    soup = BeautifulSoup(resp.text, 'html.parser')
    
    # 查找是否有我们的学生
    links = soup.select('a[data-student_id]')
    print(f"  找到 {len(links)} 个学生")
    
    # 查找 24177
    for link in links:
        student_no = link.get('data-student_no')
        if student_no == '24177':
            print(f"\n  ✓ 找到学生 24177:")
            print(f"    internal_id: {link.get('data-student_id')}")
            print(f"    name: {link.get('data-student_name')}")
            print(f"    已有成绩: {link.get('data-remark', '')}")
            
            # 检查是否已有数据
            remark = link.get_text(strip=True)
            print(f"    显示文本: {remark}")
    
    # 现在让我们直接查询一下上传后的结果
    print("\n4. 查询上次更新后的结果...")
    
    # 访问学生成绩列表
    list_url = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/index"
    resp = handler.session.get(list_url, timeout=15)
    
    # 查找最新的记录
    soup = BeautifulSoup(resp.text, 'html.parser')
    rows = soup.select('table tbody tr')
    
    print(f"  找到 {len(rows)} 条记录")
    
    # 查找包含 2026-02-06 和我们的学生的记录
    for idx, row in enumerate(rows[:5]):
        row_text = row.get_text()
        
        if '2026-02-06' in row_text or '24177' in row_text or '23121' in row_text or '23073' in row_text:
            cells = row.select('td')
            text = ' | '.join([cell.get_text(strip=True)[:40] for cell in cells[:5]])
            print(f"\n  记录 {idx}: {text}")
            
            # 查看是否有查看链接
            view_link = row.select_one('a[href*="view"]')
            if view_link:
                href = view_link.get('href')
                print(f"    链接: {href}")

except Exception as e:
    print(f"✗ 错误: {e}")
    import traceback
    traceback.print_exc()
