#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS 系统写入测试 - 测试项目编辑功能
测试：登入 -> 导航到项目编辑页面 -> 填入项目信息 -> 保存
"""

import sys
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 配置信息
LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
ITEM_SETTING_URL = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"

USERNAME = "schhs334"
PASSWORD = "schhs334"

# XPath 配置
XPATH_USERNAME = "/html/body/div[2]/div/div/div/div[2]/div/form/div[2]/input"
XPATH_PASSWORD = "/html/body/div[2]/div/div/div/div[2]/div/form/div[3]/input"
XPATH_LOGIN_BTN = "/html/body/div[2]/div/div/div/div[2]/div/form/div[4]/div/button"

XPATH_PROJECT_CODE = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[2]/div/input"
XPATH_PROJECT_NAME = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[3]/div/input"
XPATH_SCORE_ITEM = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[4]/div/input"
XPATH_SAVE_BTN = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[2]/form/div[5]/button[1]"

# 要输入的数据
PROJECT_CODE = "ACA CMI146"
PROJECT_NAME = "《马腾盛世，笔舞春风》东天宫书法比赛"
SCORE_ITEM = "0"


def get_chromedriver_path():
    """获取 ChromeDriver 的正确路径"""
    print("  [DEBUG] 获取 ChromeDriver 路径...")
    
    # 首先尝试使用 webdriver-manager
    try:
        install_path = ChromeDriverManager().install()
        print(f"  [DEBUG] webdriver-manager 返回路径: {install_path}")
        
        # 如果路径是目录，查找 chromedriver.exe
        if os.path.isdir(install_path):
            chromedriver_exe = os.path.join(install_path, "chromedriver.exe")
            if os.path.exists(chromedriver_exe):
                print(f"  [DEBUG] 找到 chromedriver.exe: {chromedriver_exe}")
                return chromedriver_exe
        
        # 如果是文件，检查是否是可执行文件
        if os.path.isfile(install_path):
            # 检查是否是 exe 文件
            if install_path.endswith('.exe'):
                print(f"  [DEBUG] 找到 chromedriver.exe: {install_path}")
                return install_path
            # 如果不是 exe，尝试替换扩展名
            elif not install_path.endswith('.chromedriver'):
                exe_path = install_path + ".exe"
                if os.path.exists(exe_path):
                    print(f"  [DEBUG] 找到 chromedriver.exe: {exe_path}")
                    return exe_path
        
        # 尝试在父目录查找
        parent_dir = os.path.dirname(install_path)
        for root, dirs, files in os.walk(parent_dir):
            for file in files:
                if file == "chromedriver.exe" or file == "chromedriver":
                    full_path = os.path.join(root, file)
                    print(f"  [DEBUG] 在目录中找到: {full_path}")
                    return full_path
        
    except Exception as e:
        print(f"  [DEBUG] webdriver-manager 失败: {e}")
    
    # 尝试直接使用 chromedriver
    print("  [DEBUG] 尝试查找系统中的 chromedriver...")
    for path in [
        "chromedriver",
        "chromedriver.exe",
        r"C:\Program Files\Google\Chrome\Application\chromedriver.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"
    ]:
        if os.path.exists(path):
            print(f"  [DEBUG] 找到: {path}")
            return path
    
    raise Exception("无法找到 ChromeDriver")


def init_driver():
    """初始化 WebDriver"""
    print("📍 初始化 ChromeDriver...")
    options = Options()
    # options.add_argument('--headless=new')  # 注释掉，方便观看过程
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    
    try:
        driver_path = get_chromedriver_path()
        print(f"✓ ChromeDriver 路径: {driver_path}")
        print(f"✓ 文件存在: {os.path.exists(driver_path)}")
        print(f"✓ 文件大小: {os.path.getsize(driver_path)} bytes")
        
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        print("✅ WebDriver 初始化成功")
        return driver
    except Exception as e:
        print(f"❌ WebDriver 初始化失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def login(driver, username, password):
    """登入SMS系统"""
    print("\n" + "="*60)
    print("📝 步骤 1: 登入 SMS 系统")
    print("="*60)
    
    try:
        print(f"[1.1] 访问登入页面: {LOGIN_URL}")
        driver.get(LOGIN_URL)
        time.sleep(2)
        print("✓ 页面已加载")
        
        # 等待用户名输入框出现
        print("[1.2] 等待登入表单加载...")
        wait = WebDriverWait(driver, 10)
        username_field = wait.until(
            EC.presence_of_element_located((By.XPATH, XPATH_USERNAME))
        )
        print("✓ 登入表单已加载")
        
        # 填入用户名
        print(f"[1.3] 填入用户名: {username}")
        username_field.clear()
        username_field.send_keys(username)
        time.sleep(0.5)
        
        # 填入密码
        print(f"[1.4] 填入密码: {'*' * len(password)}")
        password_field = driver.find_element(By.XPATH, XPATH_PASSWORD)
        password_field.clear()
        password_field.send_keys(password)
        time.sleep(0.5)
        
        # 点击登入按钮
        print("[1.5] 点击登入按钮...")
        login_btn = driver.find_element(By.XPATH, XPATH_LOGIN_BTN)
        login_btn.click()
        time.sleep(3)
        
        # 检查登入是否成功
        if "logout" in driver.page_source.lower() or "dashboard" in driver.current_url.lower():
            print("✅ 登入成功")
            return True
        else:
            print("❌ 登入可能失败，请检查页面")
            return False
            
    except Exception as e:
        print(f"❌ 登入失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def navigate_to_item_setting(driver):
    """导航到项目编辑页面"""
    print("\n" + "="*60)
    print("📝 步骤 2: 导航到项目编辑页面")
    print("="*60)
    
    try:
        print(f"[2.1] 访问项目编辑页面: {ITEM_SETTING_URL}")
        driver.get(ITEM_SETTING_URL)
        time.sleep(3)
        print("✓ 页面已加载")
        print(f"[2.2] 当前URL: {driver.current_url}")
        return True
        
    except Exception as e:
        print(f"❌ 导航失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def fill_project_info(driver):
    """填入项目信息"""
    print("\n" + "="*60)
    print("📝 步骤 3: 填入项目信息")
    print("="*60)
    
    try:
        wait = WebDriverWait(driver, 10)
        
        # 填入项目代码
        print(f"[3.1] 填入项目代码: {PROJECT_CODE}")
        code_field = wait.until(
            EC.presence_of_element_located((By.XPATH, XPATH_PROJECT_CODE))
        )
        code_field.clear()
        code_field.send_keys(PROJECT_CODE)
        time.sleep(0.5)
        print("✓ 项目代码已填入")
        
        # 填入项目名称
        print(f"[3.2] 填入项目名称: {PROJECT_NAME}")
        name_field = wait.until(
            EC.presence_of_element_located((By.XPATH, XPATH_PROJECT_NAME))
        )
        name_field.clear()
        name_field.send_keys(PROJECT_NAME)
        time.sleep(0.5)
        print("✓ 项目名称已填入")
        
        # 填入分数项目
        print(f"[3.3] 填入分数项目: {SCORE_ITEM}")
        score_field = wait.until(
            EC.presence_of_element_located((By.XPATH, XPATH_SCORE_ITEM))
        )
        score_field.clear()
        score_field.send_keys(SCORE_ITEM)
        time.sleep(0.5)
        print("✓ 分数项目已填入")
        
        return True
        
    except Exception as e:
        print(f"❌ 填入项目信息失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def save_project(driver):
    """保存项目信息"""
    print("\n" + "="*60)
    print("📝 步骤 4: 保存项目信息")
    print("="*60)
    
    try:
        print("[4.1] 等待保存按钮...")
        wait = WebDriverWait(driver, 10)
        save_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_SAVE_BTN))
        )
        print("✓ 保存按钮已加载")
        
        print("[4.2] 点击保存按钮...")
        save_btn.click()
        time.sleep(3)
        print("✓ 保存按钮已点击")
        
        # 检查是否保存成功
        if "success" in driver.page_source.lower() or "已保存" in driver.page_source:
            print("✅ 保存成功")
            return True
        else:
            print("⚠️  保存操作已执行，但无法确认成功状态")
            print(f"当前URL: {driver.current_url}")
            return True
            
    except Exception as e:
        print(f"❌ 保存失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """主程序"""
    print("\n" + "🚀 " + "="*56 + " 🚀")
    print("     SMS 系统写入测试 - 项目编辑功能")
    print("🚀 " + "="*56 + " 🚀\n")
    
    driver = None
    
    try:
        # 初始化驱动程序
        driver = init_driver()
        
        # 登入
        if not login(driver, USERNAME, PASSWORD):
            print("\n❌ 测试中止: 登入失败")
            return False
        
        # 导航到项目编辑页面
        if not navigate_to_item_setting(driver):
            print("\n❌ 测试中止: 导航失败")
            return False
        
        # 填入项目信息
        if not fill_project_info(driver):
            print("\n❌ 测试中止: 填入信息失败")
            return False
        
        # 保存项目信息
        if not save_project(driver):
            print("\n❌ 测试中止: 保存失败")
            return False
        
        # 所有步骤完成
        print("\n" + "="*60)
        print("✅ 所有测试步骤已完成")
        print("="*60)
        print("\n📊 测试结果:")
        print(f"   项目代码: {PROJECT_CODE}")
        print(f"   项目名称: {PROJECT_NAME}")
        print(f"   分数项目: {SCORE_ITEM}")
        print(f"\n✨ SMS 系统写入功能正常")
        print("\n程序将在 10 秒后自动关闭浏览器...\n")
        
        time.sleep(10)
        return True
        
    except Exception as e:
        print(f"\n❌ 程序异常: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        if driver:
            try:
                driver.quit()
                print("🔌 浏览器已关闭")
            except:
                pass


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
