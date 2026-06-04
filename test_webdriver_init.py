#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单的 WebDriver 测试 - 检查能否初始化
"""

import sys
print(f"Python: {sys.executable}")
print(f"Version: {sys.version}")
print()

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    
    print("✓ Selenium 和 webdriver-manager 已导入")
    print()
    
    # 尝试初始化 WebDriver
    print("📍 尝试初始化 ChromeDriver...")
    
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    
    try:
        driver_path = ChromeDriverManager().install()
        print(f"✓ ChromeDriver 路径: {driver_path}")
        
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        print("✅ WebDriver 初始化成功！")
        
        driver.quit()
        
    except Exception as e:
        print(f"❌ 初始化失败: {e}")
        import traceback
        traceback.print_exc()

except ImportError as e:
    print(f"❌ 导入错误: {e}")
