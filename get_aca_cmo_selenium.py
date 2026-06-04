#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
获取 ACA CMO 项目的全部资料（所有20页）- 使用 Selenium 版本
直接使用 XPath 点击分页链接
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time


def get_aca_cmo_all_pages():
    """使用 Selenium 获取 ACA CMO 项目的全部资料"""
    
    print("=" * 80)
    print("🔍 获取 ACA CMO 项目的全部资料（Selenium 版本）")
    print("=" * 80)
    print()
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    ITEM_SETTING_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"
    
    # 初始化浏览器驱动
    print("📍 第一步: 初始化浏览器驱动")
    print("-" * 80)
    
    try:
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
        print("✅ 浏览器驱动初始化成功")
    except Exception as e:
        print(f"❌ 浏览器驱动初始化失败: {e}")
        return
    
    try:
        # 2. 登入系统
        print()
        print("📍 第二步: 登入系统")
        print("-" * 80)
        
        driver.get(LOGIN_URL)
        time.sleep(2)
        
        # 等待登入表单加载
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'LoginForm_username'))
        )
        
        # 输入账号密码
        driver.find_element(By.ID, 'LoginForm_username').send_keys('schhs334')
        driver.find_element(By.ID, 'LoginForm_password').send_keys('schhs334')
        
        # 点击登入按钮
        driver.find_element(By.XPATH, "//button[@type='submit']").click()
        
        # 等待登入完成
        WebDriverWait(driver, 10).until(
            lambda d: 'login' not in d.current_url.lower()
        )
        time.sleep(1)
        print("✅ 登入成功")
        
        # 3. 访问项目设置页面
        print()
        print("📍 第三步: 访问项目设置页面")
        print("-" * 80)
        
        driver.get(ITEM_SETTING_PAGE)
        time.sleep(2)
        
        # 在搜索框中输入 ACA CMO
        print("   → 搜索关键字: ACA CMO")
        search_xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/table/thead/tr[2]/td[2]/div/input[1]"
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, search_xpath))
        )
        
        search_input = driver.find_element(By.XPATH, search_xpath)
        search_input.clear()
        search_input.send_keys("ACA CMO")
        time.sleep(1)
        
        print("✅ 搜索框输入完成")
        
        # 4. 收集所有页面的数据
        print()
        print("📍 第四步: 遍历所有 20 页")
        print("-" * 80)
        
        all_projects = []
        page_count = 0
        
        while page_count < 25:  # 最多获取25页
            page_count += 1
            print(f"⏳ 正在获取第 {page_count} 页...")
            
            # 提取当前页的表格数据
            table_xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/table/tbody/tr"
            
            try:
                rows = driver.find_elements(By.XPATH, table_xpath)
                page_projects = 0
                
                for row in rows:
                    try:
                        cells = row.find_elements(By.TAG_NAME, "td")
                        if len(cells) >= 3:
                            project_code = cells[1].text.strip()
                            project_name = cells[2].text.strip()
                            
                            if project_code and project_code not in [p['code'] for p in all_projects]:
                                all_projects.append({
                                    'code': project_code,
                                    'name': project_name
                                })
                                page_projects += 1
                    except:
                        continue
                
                print(f"   ✓ 本页找到 {page_projects} 个新项目")
            except Exception as e:
                print(f"   ⚠️  提取表格数据失败: {e}")
            
            # 尝试点击下一页
            next_page_found = False
            
            # 根据用户提供的 XPath 模式，构造下一页的 XPath
            # 第一页是 li[2], 第二页是 li[3], 等等
            next_page_li_index = page_count + 2
            
            next_page_xpath = f"/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[2]/ul/li[{next_page_li_index}]/a"
            
            try:
                next_button = driver.find_element(By.XPATH, next_page_xpath)
                # 检查按钮是否可用（不是 disabled）
                if next_button.is_enabled():
                    print(f"   → 点击第 {page_count + 1} 页链接")
                    next_button.click()
                    time.sleep(1)
                    next_page_found = True
                else:
                    print(f"   → 已到达最后一页")
                    break
            except:
                # 尝试另一种方法查找下一页
                try:
                    # 查找所有分页链接
                    pagination_links = driver.find_elements(By.XPATH, "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[2]/ul/li/a")
                    
                    # 获取当前活动页（通常有 active 类）
                    current_page_index = None
                    for i, link in enumerate(pagination_links):
                        try:
                            if 'active' in link.get_attribute('class'):
                                current_page_index = i
                                break
                        except:
                            continue
                    
                    # 点击下一个链接
                    if current_page_index is not None and current_page_index + 1 < len(pagination_links):
                        next_link = pagination_links[current_page_index + 1]
                        if next_link.is_enabled():
                            print(f"   → 点击下一页")
                            next_link.click()
                            time.sleep(1)
                            next_page_found = True
                except:
                    pass
            
            if not next_page_found:
                print(f"   → 已到达最后一页或分页导航已结束")
                break
        
        # 5. 显示结果
        print()
        print("=" * 80)
        print(f"📊 共找到 {len(all_projects)} 个 ACA CMO 项目")
        print("=" * 80)
        print()
        
        if len(all_projects) == 0:
            print("⚠️  未找到任何项目")
        else:
            # 倒序排列
            all_projects_reversed = list(reversed(all_projects))
            
            print("ACA CMO 项目列表 (倒序排列 - 最新在前):")
            print("-" * 80)
            print()
            
            for i, proj in enumerate(all_projects_reversed, 1):
                print(f"{i:4}. 代码: {proj['code']:<20} | 名称: {proj['name']}")
            
            print()
            print("=" * 80)
            print(f"✅ 共列出 {len(all_projects)} 个 ACA CMO 项目")
            print("=" * 80)
            
            # 保存到文件
            with open('ACA_CMO_全部资料.txt', 'w', encoding='utf-8') as f:
                f.write("ACA CMO 项目全部资料（倒序排列 - 最新在前）\n")
                f.write("=" * 80 + "\n\n")
                for i, proj in enumerate(all_projects_reversed, 1):
                    f.write(f"{i:4}. 代码: {proj['code']:<20} | 名称: {proj['name']}\n")
            
            print()
            print("💾 数据已保存到: ACA_CMO_全部资料.txt")
    
    except Exception as e:
        print(f"❌ 处理异常: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # 关闭浏览器
        try:
            driver.quit()
            print()
            print("✅ 浏览器已关闭")
        except:
            pass


if __name__ == "__main__":
    get_aca_cmo_all_pages()
