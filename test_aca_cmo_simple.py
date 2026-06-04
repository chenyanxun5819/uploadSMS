#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文件 - 验证 ACA CMO 项目是否能被正确查询 (简化版)
使用 requests 库直接查询，无需浏览器驱动
账号/密码: schhs334/schhs334
"""

import requests
import time
from requests.packages.urllib3.exceptions import InsecureRequestWarning

# 禁用 SSL 警告
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


def test_aca_cmo_search():
    """测试搜索 ACA CMO 项目"""
    
    print("=" * 60)
    print("🧪 ACA CMO 项目搜索测试 (轻量级版本)")
    print("=" * 60)
    print()
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    ITEM_SETTING_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"
    
    # 创建会话
    session = requests.Session()
    session.verify = False
    
    # 1. 测试连接
    print("📍 第一步: 测试连接")
    print("-" * 60)
    
    try:
        response = session.get(LOGIN_URL, timeout=10)
        if response.status_code == 200:
            print("✅ 可以访问登入页面")
        else:
            print(f"❌ 无法访问登入页面: {response.status_code}")
            return
    except Exception as e:
        print(f"❌ 连接失败: {e}")
        return
    
    print()
    time.sleep(1)
    
    # 2. 登入系统
    print("📍 第二步: 登入系统")
    print("-" * 60)
    
    login_data = {
        'LoginForm[username]': 'schhs334',
        'LoginForm[password]': 'schhs334',
        'login-button': 'login'
    }
    
    try:
        response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
        
        if 'login' not in response.url.lower():
            print(f"✅ 登入成功！")
            print(f"   当前 URL: {response.url}")
        else:
            print(f"❌ 登入失败")
            print(f"   当前 URL: {response.url}")
            return
    except Exception as e:
        print(f"❌ 登入失败: {e}")
        return
    
    print()
    time.sleep(1)
    
    # 3. 访问项目设置页面
    print("📍 第三步: 访问项目设置页面")
    print("-" * 60)
    
    try:
        response = session.get(ITEM_SETTING_PAGE, timeout=10)
        if response.status_code == 200:
            print("✅ 成功访问项目设置页面")
            
            # 检查页面中是否包含 ACA CMO
            if 'ACA CMO' in response.text:
                print("✅ 页面中找到 'ACA CMO' 文本")
            else:
                print("⚠️  页面中未找到 'ACA CMO' 文本")
            
            # 检查页面中是否包含 ACA 相关项目
            if 'ACA' in response.text:
                print("✅ 页面中找到 'ACA' 项目")
                # 查找所有 ACA 开头的项目
                import re
                aca_matches = re.findall(r'ACA[A-Za-z0-9\s]+', response.text)
                if aca_matches:
                    aca_matches = list(set(aca_matches))[:10]  # 去重并取前10个
                    print(f"   找到以下 ACA 项目:")
                    for match in aca_matches:
                        print(f"      - {match}")
            else:
                print("⚠️  页面中未找到任何 'ACA' 项目")
        else:
            print(f"❌ 无法访问项目设置页面: {response.status_code}")
            return
    except Exception as e:
        print(f"❌ 访问页面失败: {e}")
        return
    
    print()
    
    # 4. 尝试通过 API 或表单搜索项目
    print("📍 第四步: 尝试搜索项目数据")
    print("-" * 60)
    
    try:
        # 尝试搜索 ACA CMO
        search_url = ITEM_SETTING_PAGE + "&project_code=ACA CMO"
        response = session.get(search_url, timeout=10)
        
        if 'ACA CMO' in response.text:
            print("✅ 搜索 'ACA CMO' 找到结果")
        else:
            print("⚠️  搜索 'ACA CMO' 未找到结果")
        
        # 尝试搜索 ACA C
        search_url2 = ITEM_SETTING_PAGE + "&project_code=ACA C"
        response2 = session.get(search_url2, timeout=10)
        
        if 'ACA' in response2.text:
            print("✅ 搜索 'ACA C' 找到 ACA 相关项目")
        else:
            print("⚠️  搜索 'ACA C' 未找到结果")
            
    except Exception as e:
        print(f"⚠️  搜索请求异常: {e}")
    
    print()
    print("=" * 60)
    print("✅ 测试完成")
    print("=" * 60)
    print()
    print("📝 总结:")
    print("   - 如果上面出现 ✅，说明该项目在系统中存在")
    print("   - 如果出现 ⚠️，说明可能需要检查项目代码或添加项目")


if __name__ == "__main__":
    test_aca_cmo_search()
