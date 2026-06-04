#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从 SMS 页面提取班级和学会映射表
"""
import requests
from bs4 import BeautifulSoup
from pathlib import Path
import json
import datetime

USERNAME = "schhs334"
PASSWORD = "schhs334"
LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
PAGE_URL = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"

def extract_mappings():
    session = requests.Session()
    session.verify = False
    
    print("[1/3] 登入中...")
    try:
        response = session.get(LOGIN_URL, timeout=10)
        login_data = {
            'LoginForm[username]': USERNAME,
            'LoginForm[password]': PASSWORD,
            'login-button': 'login'
        }
        response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
        
        if 'login' not in response.url.lower():
            print("✓ 已登入")
        else:
            print("✗ 登入失败")
            return False
    except Exception as e:
        print(f"✗ 登入异常: {e}")
        return False
    
    print("[2/3] 获取页面...")
    try:
        response = session.get(PAGE_URL, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        print("✓ 页面已加载")
    except Exception as e:
        print(f"✗ 获取页面失败: {e}")
        return False
    
    print("[3/3] 提取数据...")
    
    # 提取班级列表
    class_select = soup.find('select', {'id': 'class_id'})
    class_mapping = {}
    if class_select:
        for option in class_select.find_all('option'):
            value = option.get('value')
            text = option.get_text(strip=True)
            if value:
                class_mapping[value] = text
        print(f"✓ 班级列表: {len(class_mapping)} 个")
    else:
        print("✗ 未找到班级选择")
    
    # 提取学会/社团列表
    club_select = soup.find('select', {'id': 'club_id'})
    club_mapping = {}
    if club_select:
        for option in club_select.find_all('option'):
            value = option.get('value')
            text = option.get_text(strip=True)
            if value:
                club_mapping[value] = text
        print(f"✓ 学会列表: {len(club_mapping)} 个")
    else:
        print("⚠ 未找到学会选择（在学生名单页面中）")
    
    # 创建目录并保存
    config_dir = Path.home() / '.sms_app'
    config_dir.mkdir(exist_ok=True)
    
    data = {
        'classes': class_mapping,
        'clubs': club_mapping,
        'timestamp': str(datetime.datetime.now())
    }
    
    config_file = config_dir / 'mappings.json'
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print(f"\n✓ 已保存到: {config_file}")
    print(f"\n班级样本 (前3个):")
    for k, v in list(class_mapping.items())[:3]:
        print(f"  {k} → {v}")
    
    print(f"\n学会样本 (前3个):")
    for k, v in list(club_mapping.items())[:3]:
        print(f"  {k} → {v}")
    
    return True

if __name__ == '__main__':
    extract_mappings()
