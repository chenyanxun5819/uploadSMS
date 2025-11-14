"""
简化登入测试：输出网页代码片段以验证是否真的登入成功
"""

import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

SMS_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
USERNAME = "schhs334"
PASSWORD = "schhs334"

def quick_login_test():
    """快速登入测试"""
    options = Options()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        print("连接到登入页面...")
        driver.get(SMS_URL)
        time.sleep(2)
        
        # 填入表单
        print("填入凭证...")
        driver.find_element(By.ID, "LoginForm_username").send_keys(USERNAME)
        driver.find_element(By.ID, "LoginForm_password").send_keys(PASSWORD)
        
        # 提交
        print("提交表单...")
        driver.find_element(By.XPATH, "//button[@type='submit']").click()
        
        # 等待并检查
        for i in range(10):
            time.sleep(1)
            url = driver.current_url
            title = driver.title
            print(f"  [{i+1}s] URL: {url}")
            
            # 检查是否离开登入页面
            if "login" not in url.lower() or i > 5:
                print(f"✓ 已离开登入页面!")
                break
        
        # 导航到活动页面
        ACTIVITY_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"
        print(f"\n导航到活动页面...")
        driver.get(ACTIVITY_PAGE)
        time.sleep(3)
        
        url = driver.current_url
        title = driver.title
        print(f"✓ 页面加载完成!")
        print(f"  URL: {url}")
        print(f"  标题: {title}")
        page_source = driver.page_source
        
        # 检查是否有活动表单
        if "StudentPerformanceM_item_id" in page_source:
            print("✓ 找到活动输入框标记!")
            # 找出包含此标记的行
            for line in page_source.split('\n'):
                if "StudentPerformanceM_item_id" in line:
                    print(f"  {line[:120]}")
        else:
            print("✗ 未找到活动输入框标记")
            print("\n页面上前2000个字符:")
            print(page_source[:2000])
        
        return True
    
    except Exception as e:
        print(f"✗ 错误: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        driver.quit()

if __name__ == "__main__":
    quick_login_test()
