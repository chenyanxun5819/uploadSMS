#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单的 SMS 登入测试 - 直接测试连接
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

def test_login(username, password):
    """测试 SMS 登入"""
    print("=" * 60)
    print(f"🧪 SMS 登入测试")
    print("=" * 60)
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    
    try:
        # 初始化浏览器
        print("\n[1] 初始化浏览器...")
        options = Options()
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
        driver.set_window_size(1200, 900)
        print("✓ 浏览器已初始化")
        
        # 打开登入页面
        print(f"\n[2] 打开登入页面: {LOGIN_URL}")
        driver.get(LOGIN_URL)
        print("✓ 页面已加载")
        time.sleep(2)
        
        # 等待登入表单
        print(f"\n[3] 等待登入表单...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'LoginForm_username'))
        )
        print("✓ 登入表单已找到")
        
        # 输入帐号
        print(f"\n[4] 输入帐号: {username}")
        username_field = driver.find_element(By.ID, 'LoginForm_username')
        username_field.clear()
        username_field.send_keys(username)
        print("✓ 帐号已输入")
        time.sleep(1)
        
        # 输入密码
        print(f"\n[5] 输入密码: {'*' * len(password)}")
        password_field = driver.find_element(By.ID, 'LoginForm_password')
        password_field.clear()
        password_field.send_keys(password)
        print("✓ 密码已输入")
        time.sleep(1)
        
        # 点击登入按钮
        print(f"\n[6] 点击登入按钮...")
        submit_btn = driver.find_element(By.XPATH, "//button[@type='submit']")
        submit_btn.click()
        print("✓ 按钮已点击")
        time.sleep(3)
        
        # 等待登入完成
        print(f"\n[7] 等待登入完成...")
        WebDriverWait(driver, 15).until(
            lambda d: 'login' not in d.current_url.lower()
        )
        time.sleep(2)
        
        # 检查结果
        print(f"\n[8] 检查登入结果...")
        current_url = driver.current_url
        print(f"✓ 当前 URL: {current_url}")
        
        if 'login' not in current_url.lower():
            print("\n" + "=" * 60)
            print("✅ 登入成功！")
            print("=" * 60)
            driver.quit()
            return True
        else:
            print("\n" + "=" * 60)
            print("❌ 登入失败 - 仍在登入页面")
            print("=" * 60)
            driver.quit()
            return False
            
    except Exception as e:
        print(f"\n❌ 异常发生: {e}")
        print("=" * 60)
        try:
            driver.quit()
        except:
            pass
        return False

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        print("使用方法: python test_login.py <帐号> <密码>")
        print("示例: python test_login.py schhs334 schhs334")
        sys.exit(1)
    
    username = sys.argv[1]
    password = sys.argv[2]
    
    result = test_login(username, password)
    sys.exit(0 if result else 1)
