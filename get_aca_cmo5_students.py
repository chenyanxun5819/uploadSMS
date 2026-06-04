#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
获取 ACA CMO5 项目的学生成绩数据（所有页面）
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


def get_aca_cmo5_student_scores():
    """获取 ACA CMO5 的学生成绩（所有页面）"""
    
    print("=" * 90)
    print("🎓 获取 ACA CMO5 项目的学生成绩数据（所有页面）")
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
    
    # 2. 访问成绩输入页面，并过滤 ACA CMO5
    print("📍 第二步: 搜索 ACA CMO5 学生成绩")
    print("-" * 90)
    
    all_students = []
    page = 0
    
    for page_num in range(1, 50):  # 最多获取 50 页
        page_num_display = page_num
        print(f"⏳ 正在获取第 {page_num_display} 页...")
        
        try:
            # 构造搜索参数
            search_params = {
                'StudentPerformance[mark_item]': 'ACA CMO5',  # 搜索项目代码
                'page': page_num
            }
            
            response = session.get(PERFORMANCE_PAGE, params=search_params, timeout=10)
            
            if response.status_code != 200:
                print(f"   ❌ 页面加载失败")
                break
            
            # 解析学生表格
            parser = StudentTableParser()
            parser.feed(response.text)
            rows = parser.rows
            
            if not rows or len(rows) < 1:
                print(f"   → 已到达最后一页（无数据）")
                break
            
            # 提取学生数据
            page_students = 0
            for row in rows:
                if len(row) >= 4:
                    try:
                        # 假设表格结构: 学号, 姓名, 性别, 成绩等
                        student_id = row[0].strip() if len(row) > 0 else ""
                        student_name = row[1].strip() if len(row) > 1 else ""
                        score = row[3].strip() if len(row) > 3 else ""
                        
                        if student_id and student_name:
                            all_students.append({
                                'id': student_id,
                                'name': student_name,
                                'score': score,
                                'page': page_num_display
                            })
                            page_students += 1
                    except:
                        continue
            
            if page_students > 0:
                print(f"   ✓ 本页找到 {page_students} 个学生")
                page = page_num_display
            else:
                print(f"   → 本页无学生数据")
                break
            
        except Exception as e:
            print(f"   ⚠️  第 {page_num_display} 页加载异常: {e}")
            if page_num == 1:
                break
        
        time.sleep(0.5)
    
    # 3. 显示结果
    print()
    print("=" * 90)
    print(f"📊 共获取 {page} 页数据，找到 {len(all_students)} 个学生")
    print("=" * 90)
    print()
    
    if len(all_students) == 0:
        print("⚠️  未找到任何学生数据")
        # 尝试查看第一页的原始内容
        print("\n📝 尝试获取第一页的原始数据以调试...")
        try:
            response = session.get(PERFORMANCE_PAGE + "?mark_item=ACA%20CMO5", timeout=10)
            with open('performance_page1.html', 'w', encoding='utf-8') as f:
                f.write(response.text)
            print("   已保存原始页面到: performance_page1.html")
        except:
            pass
        return
    
    # 倒序排列
    all_students_reversed = list(reversed(all_students))
    
    print("ACA CMO5 学生成绩列表 (倒序排列 - 最新在前):")
    print("-" * 90)
    print(f"{'序号':<6} {'学号':<12} {'姓名':<20} {'成绩':<12} {'页码':<6}")
    print("-" * 90)
    
    for i, student in enumerate(all_students_reversed, 1):
        print(f"{i:<6} {student['id']:<12} {student['name']:<20} {student['score']:<12} {student['page']:<6}")
    
    print()
    print("=" * 90)
    print(f"✅ 共列出 {len(all_students)} 个学生记录")
    print("=" * 90)
    
    # 保存到文件
    with open('ACA_CMO5_学生成绩.txt', 'w', encoding='utf-8') as f:
        f.write("ACA CMO5 项目学生成绩数据（倒序排列 - 最新在前）\n")
        f.write("=" * 90 + "\n\n")
        f.write(f"{'序号':<6} {'学号':<12} {'姓名':<20} {'成绩':<12} {'页码':<6}\n")
        f.write("-" * 90 + "\n")
        for i, student in enumerate(all_students_reversed, 1):
            f.write(f"{i:<6} {student['id']:<12} {student['name']:<20} {student['score']:<12} {student['page']:<6}\n")
    
    print()
    print("💾 数据已保存到: ACA_CMO5_学生成绩.txt")


if __name__ == "__main__":
    get_aca_cmo5_student_scores()
