"""
测试脚本：选择班级并依据 studentId 查找学生资料
用法示例：
  python test_select_class_and_find_student.py --class S3B --student 20019

脚本步骤：
- 登入（请在代码中配置或使用环境变数）
- 导航到活动创建页并确保已填入日期与活动
- 点击学生名单按钮（id= yw4），选择班级并在表格中寻找学号

注意：此脚本为示例，凭证建议使用环境变量或受保护的配置
"""

import argparse
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 配置（改用环境变数更安全）
SMS_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
USERNAME = os.getenv('SMS_USERNAME', 'schhs334')
PASSWORD = os.getenv('SMS_PASSWORD', 'schhs334')
ACTIVITY_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/create"


def login_to_sms(driver, timeout=10):
    driver.get(SMS_URL)
    WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, 'LoginForm_username')))
    driver.find_element(By.ID, 'LoginForm_username').clear()
    driver.find_element(By.ID, 'LoginForm_username').send_keys(USERNAME)
    driver.find_element(By.ID, 'LoginForm_password').clear()
    driver.find_element(By.ID, 'LoginForm_password').send_keys(PASSWORD)
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    WebDriverWait(driver, timeout).until(lambda d: 'login' not in d.current_url.lower())
    time.sleep(1)
    print('✓ 登入成功')


def select_class_and_find_student(driver, class_short: str, student_id: str, timeout=8):
    # 假設已在活动页面
    # 点击学生名单按钮（若页面未显示，可先调用 driver.find_element(By.ID, 'yw4').click()）
    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, 'yw4'))).click()
    except Exception:
        # 若找不到或不可点，忽略（有些页面可能已自动展开）
        pass

    sel = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, 'class_id'))
    )
    select = Select(sel)

    matched_value = None
    for opt in sel.find_elements(By.TAG_NAME, 'option'):
        text = opt.text.strip()
        if f'({class_short})' in text or text.endswith(class_short):
            matched_value = opt.get_attribute('value')
            break

    if not matched_value:
        print(f"未找到匹配的班級選項: {class_short}")
        return None

    select.select_by_value(matched_value)

    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'table.table tbody tr'))
        )
    except Exception:
        print('表格未在時間內載入完畢')
        return None

    rows = driver.find_elements(By.CSS_SELECTOR, 'table.table tbody tr')
    for r in rows:
        cols = r.find_elements(By.TAG_NAME, 'td')
        if not cols:
            continue
        stu_no = cols[0].text.strip()
        if stu_no == str(student_id):
            return {
                'student_no': stu_no,
                'name_en': cols[1].text.strip() if len(cols) > 1 else '',
                'name_cn': cols[2].text.strip() if len(cols) > 2 else '',
                'class_name': cols[3].text.strip() if len(cols) > 3 else '',
            }

    print(f'在表格中未找到學號 {student_id}')
    return None


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--class', dest='class_short', required=True, help='班級簡寫，例如 S3B')
    parser.add_argument('--student', dest='student_id', required=True, help='要查找的学号，例如 20019')
    args = parser.parse_args()

    options = Options()
    options.add_argument('--disable-blink-features=AutomationControlled')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    try:
        login_to_sms(driver)
        driver.get(ACTIVITY_PAGE)
        time.sleep(1)

        res = select_class_and_find_student(driver, args.class_short, args.student_id)
        if res:
            print('找到學生：', res)
        else:
            print('找不到學生')
    finally:
        driver.quit()