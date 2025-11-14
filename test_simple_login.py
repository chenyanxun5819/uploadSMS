"""
简化版测试：仅验证登入成功，然后尝试抓取页面元素（用于调试 XPath）
"""

import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import load_workbook

# SMS 系统凭证
SMS_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
USERNAME = os.getenv("SMS_USERNAME", "schhs334")
PASSWORD = os.getenv("SMS_PASSWORD", "schhs334")

# 登入表单 XPath（使用通用 ID 选择器）
USERNAME_XPATH = '//*[@id="LoginForm_username"]'
PASSWORD_XPATH = '//*[@id="LoginForm_password"]'
SUBMIT_XPATH = '//button[@type="submit"]'

# 活动页面
ACTIVITY_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"

# Excel 文件路径
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "Upload.xlsx")

def test_login():
    """
    测试登入功能，并返回登入后的页面标题或 URL
    """
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        print(f"[1] 正在连接到 {SMS_URL}...")
        driver.get(SMS_URL)
        time.sleep(2)
        
        print("[2] 正在填入凭证...")
        username_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, USERNAME_XPATH))
        )
        username_field.clear()
        username_field.send_keys(USERNAME)
        print(f"   ✓ 输入用户名: {USERNAME}")
        time.sleep(0.5)
        
        password_field = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, PASSWORD_XPATH))
        )
        password_field.clear()
        password_field.send_keys(PASSWORD)
        print("   ✓ 输入密码")
        time.sleep(0.5)
        
        print("[3] 点击登入按钮...")
        submit_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, SUBMIT_XPATH))
        )
        submit_btn.click()
        
        # 等待页面重定向或加载完成（等待登入页面消失）
        print("      等待登入重定向...")
        try:
            WebDriverWait(driver, 10).until(
                EC.url_changes(SMS_URL)
            )
            print("      ✓ 页面已重定向")
        except Exception as e:
            print(f"      ⚠ 页面未重定向，可能需要额外等待: {e}")
        
        time.sleep(3)
        
        current_url = driver.current_url
        page_title = driver.title
        print(f"✓ 登入成功！")
        print(f"   当前 URL: {current_url}")
        print(f"   页面标题: {page_title}")
        
        # 导航到活动页面
        print(f"\n[4] 导航到活动页面...")
        driver.get(ACTIVITY_PAGE)
        time.sleep(3)
        
        current_url = driver.current_url
        page_title = driver.title
        print(f"✓ 页面加载完成！")
        print(f"   当前 URL: {current_url}")
        print(f"   页面标题: {page_title}")
        
        # 尝试找到活动下拉菜单的不同选择器
        print(f"\n[5] 尝试查找活动下拉菜单元素...")
        
        possible_xpaths = [
            '//*[@id="s2id_StudentPerformanceM_item_id"]/a',
            '//input[@id="StudentPerformanceM_item_id"]',
            '//div[@id="s2id_StudentPerformanceM_item_id"]',
            '//span[contains(@id, "StudentPerformanceM_item_id")]',
        ]
        
        for i, xpath in enumerate(possible_xpaths):
            try:
                element = driver.find_element(By.XPATH, xpath)
                print(f"   ✓ XPath {i+1} 找到元素: {xpath}")
                print(f"      标签: {element.tag_name}, 文本: {element.text[:50] if element.text else '(无文本)'}")
            except Exception as e:
                print(f"   ✗ XPath {i+1} 未找到: {xpath}")
        
        # 尝试获取所有输入框
        print(f"\n[6] 页面上所有的输入框或 select 元素:")
        try:
            inputs = driver.find_elements(By.TAG_NAME, "input")
            selects = driver.find_elements(By.TAG_NAME, "select")
            print(f"   找到 {len(inputs)} 个 input 元素")
            for inp in inputs[:5]:  # 仅显示前5个
                print(f"      - ID: {inp.get_attribute('id')}, Name: {inp.get_attribute('name')}, Type: {inp.get_attribute('type')}")
            print(f"   找到 {len(selects)} 个 select 元素")
            for sel in selects[:5]:
                print(f"      - ID: {sel.get_attribute('id')}, Name: {sel.get_attribute('name')}")
        except Exception as e:
            print(f"   ✗ 无法获取输入元素: {e}")
        
        return True
    
    except Exception as e:
        print(f"✗ 错误: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        print("\n[X] 关闭浏览器...")
        driver.quit()

if __name__ == "__main__":
    success = test_login()
    if success:
        print("\n✓ 测试完成！登入成功，可以开始实装完整脚本。")
    else:
        print("\n✗ 测试失败，请检查凭证或网页结构。")
