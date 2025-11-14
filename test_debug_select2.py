"""
深度调试：输出 Select2 打开后的 DOM 结构
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
ACTIVITY_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"

options = Options()
options.add_argument('--disable-blink-features=AutomationControlled')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    # 登入
    print("登入...")
    driver.get(SMS_URL)
    time.sleep(2)
    driver.find_element(By.ID, "LoginForm_username").send_keys(USERNAME)
    driver.find_element(By.ID, "LoginForm_password").send_keys(PASSWORD)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    WebDriverWait(driver, 10).until(lambda d: "login" not in d.current_url.lower())
    time.sleep(2)
    print("✓ 登入成功")
    
    # 导航到活动页面
    print("导航到活动页面...")
    driver.get(ACTIVITY_PAGE)
    time.sleep(3)
    
    # 点击 Select2
    print("点击 Select2...")
    select2 = driver.find_element(By.ID, "s2id_StudentPerformanceM_item_id")
    select2.click()
    time.sleep(2)
    
    # 输入搜索文本
    print("输入搜索文本...")
    search_box = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//input"))
    )
    search_box.send_keys("ACA CMO183")
    time.sleep(2)
    
    # 获取 DOM 片段
    print("\n=== Select2 打开后的 DOM 结构 ===\n")
    
    # 输出搜索框周围的 DOM
    search_container_html = driver.execute_script("""
    var container = document.querySelector('.select2-input').closest('.select2-container');
    if (container) {
        return container.outerHTML.substring(0, 1500);
    }
    return "未找到容器";
    """)
    print("搜索框容器 HTML:")
    print(search_container_html)
    
    # 查找所有 select2-result-label
    results_html = driver.execute_script("""
    var results = document.querySelectorAll('.select2-result-label');
    var output = [];
    results.forEach(function(el) {
        output.push(el.innerText);
    });
    return output.length > 0 ? output : "未找到结果";
    """)
    print(f"\n找到的 select2-result-label: {results_html}")
    
    # 查找所有下拉选项（可能的 class）
    dropdown_html = driver.execute_script("""
    var possible_selectors = [
        '.select2-results',
        '.select2-result',
        '[class*="select2-result"]',
        '.select2-drop',
        '.select2-dropdown'
    ];
    
    var output = {};
    possible_selectors.forEach(function(selector) {
        var elements = document.querySelectorAll(selector);
        output[selector] = elements.length;
    });
    
    return output;
    """)
    print(f"\n可能的选择器及数量: {dropdown_html}")
    
    # 输出整个可见的下拉菜单 HTML（前 2000 字符）
    dropdown_content = driver.execute_script("""
    var dropdown = document.querySelector('.select2-drop');
    if (dropdown) {
        return dropdown.outerHTML.substring(0, 2000);
    }
    return "未找到 .select2-drop";
    """)
    print(f"\n.select2-drop HTML（前 2000 字符）:\n{dropdown_content}")
    
finally:
    driver.quit()
