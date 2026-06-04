#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动检测最后一页，然后获取 ACA CMO5 项目的全部学生成绩数据
不需要预先告诉我们共有多少页
"""

import requests
import time
from html.parser import HTMLParser
from requests.packages.urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)


class StudentTableParser(HTMLParser):
    """学生成绩表格解析器"""
    def __init__(self):
        super().__init__()
        self.rows = []
        self.current_row = []
        self.in_tbody = False
        self.in_td = False
        self.current_cell = ""
    
    def handle_starttag(self, tag, attrs):
        if tag == "tbody":
            self.in_tbody = True
        elif tag == "tr" and self.in_tbody:
            self.current_row = []
        elif tag in ["td", "th"] and self.in_tbody:
            self.in_td = True
            self.current_cell = ""
    
    def handle_endtag(self, tag):
        if tag == "tbody":
            self.in_tbody = False
        elif tag == "tr" and self.in_tbody:
            if self.current_row:
                self.rows.append(self.current_row)
        elif tag in ["td", "th"] and self.in_tbody:
            self.in_td = False
            self.current_row.append(self.current_cell.strip())
    
    def handle_data(self, data):
        if self.in_td:
            self.current_cell += data


def check_page_has_data(session, url, page_num):
    """检查某一页是否有数据"""
    try:
        search_params = {
            'StudentPerformance[mark_item]': 'ACA CMO5',
            'page': page_num
        }
        response = session.get(url, params=search_params, timeout=10)
        
        if response.status_code != 200:
            return False
        
        parser = StudentTableParser()
        parser.feed(response.text)
        return len(parser.rows) > 0
    except:
        return False


def find_last_page(session, url, max_search=1000):
    """
    使用二分查找法自动找出最后一页
    
    策略:
    1. 先检查一个较大的页码（比如1000）
    2. 如果没数据，逐页往回找
    3. 找到第一个有数据的页面，那就是最后一页
    """
    print("🔍 自动检测最后一页...")
    print("-" * 90)
    
    # 先尝试一个大页码
    print(f"⏳ 测试第 {max_search} 页...", end="", flush=True)
    if check_page_has_data(session, url, max_search):
        print(" ✓ 有数据")
        # 如果有数据，说明最后一页可能比1000还大
        print("⚠️  数据超出预期范围，扩大搜索...")
        for page in range(max_search + 100, max_search + 10000, 100):
            if not check_page_has_data(session, url, page):
                # 找到了上限，现在在这个范围内二分查找
                return binary_search_last_page(session, url, max_search, page)
        return max_search
    else:
        print(" ✗ 无数据")
    
    # 逐页往回找第一个有数据的页面
    print(f"⏳ 从第 {max_search} 页往回查找第一个有数据的页面...")
    
    for page in range(max_search - 1, 0, -1):
        if page % 50 == 0:
            print(f"  → 检查第 {page} 页...", end="", flush=True)
        
        if check_page_has_data(session, url, page):
            if page % 50 != 0:
                print()
            print(f"✅ 找到最后一页: 第 {page} 页")
            return page
        
        if page % 50 == 0:
            print(" ✗", end="")
        
        time.sleep(0.1)
    
    print()
    print("❌ 找不到任何有数据的页面")
    return 0


def binary_search_last_page(session, url, low, high):
    """二分查找最后一页"""
    print(f"📊 在第 {low} 页到 {high} 页之间进行二分查找...")
    
    last_page_with_data = low
    
    while low <= high:
        mid = (low + high) // 2
        if check_page_has_data(session, url, mid):
            last_page_with_data = mid
            low = mid + 1
            print(f"  ✓ 第 {mid} 页有数据，继续往后查找...")
        else:
            high = mid - 1
            print(f"  ✗ 第 {mid} 页无数据，往前查找...")
        time.sleep(0.2)
    
    print(f"✅ 最后一页是: 第 {last_page_with_data} 页")
    return last_page_with_data


def get_aca_cmo5_student_scores():
    """获取 ACA CMO5 的学生成绩（自动检测所有页面）"""
    
    print("=" * 90)
    print("🎓 获取 ACA CMO5 项目的学生成绩数据（自动检测最后一页）")
    print("=" * 90)
    print()
    
    LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
    PERFORMANCE_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/studentPerformance/index"
    
    session = requests.Session()
    session.verify = False
    
    # 1. 登入
    print("📍 第一步: 登入系统")
    print("-" * 90)
    
    login_data = {
        'LoginForm[username]': 'schhs334',
        'LoginForm[password]': 'schhs334',
        'login-button': 'login'
    }
    
    try:
        response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
        if 'login' not in response.url.lower():
            print("✅ 登入成功")
        else:
            print("❌ 登入失败")
            return
    except Exception as e:
        print(f"❌ 登入失败: {e}")
        return
    
    print()
    time.sleep(1)
    
    # 2. 自动检测最后一页
    print("📍 第二步: 自动检测最后一页")
    print("-" * 90)
    
    last_page = find_last_page(session, PERFORMANCE_PAGE, max_search=500)
    
    if last_page == 0:
        print("❌ 无法找到有效的数据页面")
        return
    
    print()
    
    # 3. 获取所有页的数据
    print("📍 第三步: 获取第 1 到 第 {0} 页的所有数据".format(last_page))
    print("-" * 90)
    
    all_students = []
    
    for page_num in range(1, last_page + 1):
        print(f"⏳ 正在获取第 {page_num}/{last_page} 页...", end="", flush=True)
        
        try:
            search_params = {
                'StudentPerformance[mark_item]': 'ACA CMO5',
                'page': page_num
            }
            
            response = session.get(PERFORMANCE_PAGE, params=search_params, timeout=10)
            
            if response.status_code != 200:
                print(" ❌ 页面加载失败")
                continue
            
            parser = StudentTableParser()
            parser.feed(response.text)
            rows = parser.rows
            
            page_students = 0
            for row in rows:
                if len(row) >= 4:
                    try:
                        student_id = row[0].strip() if len(row) > 0 else ""
                        student_name = row[1].strip() if len(row) > 1 else ""
                        score = row[3].strip() if len(row) > 3 else ""
                        
                        if student_id and student_name:
                            all_students.append({
                                'id': student_id,
                                'name': student_name,
                                'score': score,
                                'page': page_num
                            })
                            page_students += 1
                    except:
                        continue
            
            print(f" ✓ {page_students} 个学生")
            
        except Exception as e:
            print(f" ⚠️  异常: {e}")
        
        time.sleep(0.3)
    
    # 4. 显示结果
    print()
    print("=" * 90)
    print(f"📊 自动检测结果: 共 {last_page} 页，找到 {len(all_students)} 条记录")
    print("=" * 90)
    print()
    
    if len(all_students) == 0:
        print("⚠️  未找到任何学生数据")
        return
    
    # 倒序排列
    all_students_reversed = list(reversed(all_students))
    
    print("ACA CMO5 学生成绩列表 (倒序排列 - 最新在前):")
    print("-" * 90)
    print(f"{'序号':<6} {'学号':<12} {'姓名':<20} {'成绩':<12} {'页码':<6}")
    print("-" * 90)
    
    for i, student in enumerate(all_students_reversed[:20], 1):  # 只显示前20条
        print(f"{i:<6} {student['id']:<12} {student['name']:<20} {student['score']:<12} {student['page']:<6}")
    
    if len(all_students) > 20:
        print(f"... 共 {len(all_students)} 条记录 ...")
    
    print()
    
    # 保存到文件
    output_file = 'ACA_CMO5_学生成绩_自动检测.txt'
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("ACA CMO5 项目学生成绩数据（自动检测最后一页 - 倒序排列）\n")
        f.write("=" * 90 + "\n\n")
        f.write(f"自动检测结果: 共 {last_page} 页\n\n")
        f.write(f"{'序号':<6} {'学号':<12} {'姓名':<20} {'成绩':<12} {'页码':<6}\n")
        f.write("-" * 90 + "\n")
        for i, student in enumerate(all_students_reversed, 1):
            f.write(f"{i:<6} {student['id']:<12} {student['name']:<20} {student['score']:<12} {student['page']:<6}\n")
    
    print("=" * 90)
    print(f"✅ 已保存到文件: {output_file}")
    print("=" * 90)


if __name__ == '__main__':
    get_aca_cmo5_student_scores()
