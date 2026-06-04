#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试脚本 - 检查 HTML 结构
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "sms_app" / "core"))

from sms_handler_requests_v1 import SMSHandlerRequests
from bs4 import BeautifulSoup

username = "schhs334"
password = "schhs334"

handler = SMSHandlerRequests()

# 登入
print(">>> Logging in...")
if not handler.login(username, password):
    print("FAILED to login")
    sys.exit(1)

print(">>> Fetching activity page...")
resp = handler.session.get(handler.ACTIVITY_PAGE, timeout=15)
print(f"Status: {resp.status_code}")

# 保存 HTML
with open('debug_activity_page.html', 'w', encoding='utf-8') as f:
    f.write(resp.text)

print(">>> Saved HTML to debug_activity_page.html")

# 解析
soup = BeautifulSoup(resp.text, 'html.parser')

# 查找所有学生链接
student_links = soup.select('a[onclick="addToEkstra(this)"]')
print(f">>> Found {len(student_links)} student links with a[onclick=addToEkstra]")

# 查找所有包含 data-student_id 的元素
data_student = soup.find_all(attrs={'data-student_id': True})
print(f">>> Found {len(data_student)} elements with data-student_id")

# 查找所有 a 标签
all_links = soup.find_all('a')
print(f">>> Found {len(all_links)} total links")

# 显示前 10 个链接的详情
print("\n>>> First 10 links:")
for i, link in enumerate(all_links[:10]):
    attrs = dict(link.attrs)
    print(f"{i}: {link.name} - {attrs}")

# 查找模态框
modals = soup.find_all(attrs={'class': 'modal'})
print(f"\n>>> Found {len(modals)} modals")

# 查找表格
tables = soup.find_all('table')
print(f">>> Found {len(tables)} tables")

# 查找 div 中的数据
grids = soup.find_all(id='StudentPerformanceM_grid')
print(f">>> Found {len(grids)} grids")

# 查找学生列表选择器
student_lists = soup.select('[id*="student"]')
print(f">>> Found {len(student_lists)} student-related elements")

# 保存测试结果
print("\n>>> Analysis complete. Check debug_activity_page.html for details.")
