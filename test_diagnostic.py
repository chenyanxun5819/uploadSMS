#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS 诊断脚本 - 比对 test_sms_write.py 和 project_input_page.py 的行为
检查登入错误和项目数据提交问题
"""

import sys
import time
sys.path.insert(0, 'sms_app')

from core.sms_handler import SMSHandler
from core.config_manager import ConfigManager

USERNAME = "schhs334"
PASSWORD = "schhs334"
PROJECT_CODE = "ACA CMI146"
PROJECT_NAME = "《马腾盛世，笔舞春风》东天宫书法比赛"
SCORE_ITEM = "0"


def test_direct_selenium():
    """测试 1: 直接使用 Selenium（参考 test_sms_write.py）"""
    print("\n" + "="*70)
    print("📝 测试 1: 直接使用 Selenium（参考 test_sms_write.py 的方法）")
    print("="*70)
    
    handler = SMSHandler(headless=False)
    
    try:
        print("\n[1.1] 初始化驱动...")
        if not handler.init_driver():
            print("❌ 驱动初始化失败")
            return False
        print("✅ 驱动已初始化")
        
        print("\n[1.2] 登入系统...")
        if not handler.login(USERNAME, PASSWORD, timeout=15):
            print("❌ 登入失败")
            return False
        print("✅ 登入成功")
        
        print("\n[1.3] 添加项目...")
        print(f"   项目代码: {PROJECT_CODE}")
        print(f"   项目名称: {PROJECT_NAME}")
        print(f"   分数项目: {SCORE_ITEM}")
        
        if not handler.add_project(PROJECT_CODE, PROJECT_NAME, SCORE_ITEM):
            print("❌ 项目添加失败")
            return False
        print("✅ 项目已添加")
        
        print("\n✅ 测试 1 通过")
        return True
        
    except Exception as e:
        print(f"\n❌ 异常: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        handler.close_driver()


def test_project_input_page_style():
    """测试 2: 使用 project_input_page.py 的方式（带线程）"""
    print("\n" + "="*70)
    print("📝 测试 2: 使用 project_input_page.py 的方式（AddProjectThread）")
    print("="*70)
    
    from sms_app.ui.pages.project_input_page import AddProjectThread
    
    try:
        print("\n[2.1] 创建线程...")
        thread = AddProjectThread(USERNAME, PASSWORD, PROJECT_CODE, PROJECT_NAME, SCORE_ITEM)
        print("✅ 线程已创建")
        
        # 连接信号
        success_flag = {'success': False, 'message': ''}
        
        def on_finished(success, message):
            success_flag['success'] = success
            success_flag['message'] = message
        
        thread.add_finished.connect(on_finished)
        
        print("\n[2.2] 启动线程...")
        thread.start()
        
        # 等待线程完成
        print("[2.3] 等待线程完成（最多 60 秒）...")
        thread.wait(60000)  # 60 秒超时
        
        print(f"\n✅ 线程完成: {success_flag['message']}")
        
        if success_flag['success']:
            print("✅ 测试 2 通过")
            return True
        else:
            print("❌ 测试 2 失败")
            return False
            
    except Exception as e:
        print(f"\n❌ 异常: {e}")
        import traceback
        traceback.print_exc()
        return False


def compare_xpath_values():
    """测试 3: 比对 XPath 值和数据"""
    print("\n" + "="*70)
    print("📝 测试 3: 比对 XPath 值和数据完整性")
    print("="*70)
    
    handler = SMSHandler(headless=False)
    
    try:
        print("\n[3.1] 初始化并登入...")
        handler.init_driver()
        handler.login(USERNAME, PASSWORD, timeout=15)
        
        print("\n[3.2] 导航到项目编辑页面...")
        handler.driver.get("http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index")
        time.sleep(2)
        
        print("\n[3.3] 检查各个输入框的信息...")
        
        # 检查项目代码框
        code_xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[2]/div/input"
        code_input = handler.driver.find_element("xpath", code_xpath)
        print(f"\n✅ 项目代码输入框:")
        print(f"   - 选择器有效: ✓")
        print(f"   - placeholder: {code_input.get_attribute('placeholder')}")
        print(f"   - 当前值: '{code_input.get_attribute('value')}'")
        print(f"   - 测试数据: {PROJECT_CODE}")
        
        # 检查项目名称框
        name_xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[3]/div/input"
        name_input = handler.driver.find_element("xpath", name_xpath)
        print(f"\n✅ 项目名称输入框:")
        print(f"   - 选择器有效: ✓")
        print(f"   - placeholder: {name_input.get_attribute('placeholder')}")
        print(f"   - 当前值: '{name_input.get_attribute('value')}'")
        print(f"   - 测试数据: {PROJECT_NAME}")
        
        # 检查分数框
        score_xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[4]/div/input"
        score_input = handler.driver.find_element("xpath", score_xpath)
        print(f"\n✅ 分数项目输入框:")
        print(f"   - 选择器有效: ✓")
        print(f"   - placeholder: {score_input.get_attribute('placeholder')}")
        print(f"   - 当前值: '{score_input.get_attribute('value')}'")
        print(f"   - 测试数据: {SCORE_ITEM}")
        
        print("\n✅ 测试 3 通过")
        return True
        
    except Exception as e:
        print(f"\n❌ 异常: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        handler.close_driver()


def main():
    print("\n🚀 " + "="*66 + " 🚀")
    print("   SMS 系统诊断 - 比对不同实现方法和检查登入错误")
    print("🚀 " + "="*66 + " 🚀")
    
    results = {
        '测试1-直接Selenium': False,
        '测试2-项目页面线程': False,
        '测试3-XPath数据': False
    }
    
    try:
        # 测试 1
        results['测试1-直接Selenium'] = test_direct_selenium()
        
        # 测试 2
        results['测试2-项目页面线程'] = test_project_input_page_style()
        
        # 测试 3
        results['测试3-XPath数据'] = compare_xpath_values()
        
    except Exception as e:
        print(f"\n❌ 主程序异常: {e}")
        import traceback
        traceback.print_exc()
    
    # 输出总结
    print("\n" + "="*70)
    print("📊 测试总结")
    print("="*70)
    
    for test_name, result in results.items():
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{test_name}: {status}")
    
    all_passed = all(results.values())
    
    if all_passed:
        print("\n✅ 所有测试通过 - SMS 系统工作正常")
    else:
        print("\n❌ 部分测试失败 - 请检查错误日志")
    
    return 0 if all_passed else 1


if __name__ == "__main__":
    sys.exit(main())
