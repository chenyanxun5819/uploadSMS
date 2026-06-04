#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试脚本 - 查看搜索结果的详细内容
"""

import requests
import time
from requests.packages.urllib3.exceptions import InsecureRequestWarning

requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

LOGIN_URL = "http://sms.chhsban.edu.my/sms/index.php?r=site/login"
ITEM_SETTING_PAGE = "http://sms.chhsban.edu.my/sms/index.php?r=transaction/itemSetting/index"

# 登入
session = requests.Session()
session.verify = False

login_data = {
    'LoginForm[username]': 'schhs334',
    'LoginForm[password]': 'schhs334',
    'login-button': 'login'
}

response = session.post(LOGIN_URL, data=login_data, timeout=10, allow_redirects=True)
print("✅ 登入成功\n")

# 搜索 ACA CMO
search_params = {
    'ItemSetting[item_code]': 'ACA CMO',
}

response = session.get(ITEM_SETTING_PAGE, params=search_params, timeout=10)

# 保存页面内容到文件，方便查看
with open('search_result_page1.html', 'w', encoding='utf-8') as f:
    f.write(response.text)

print("📄 第一页 HTML 已保存到: search_result_page1.html")
print()

# 分析页面中的关键内容
html = response.text

# 查找分页信息
import re

# 查找分页导航
pagination_match = re.search(r'<ul[^>]*class="pagination"[^>]*>.*?</ul>', html, re.DOTALL)
if pagination_match:
    pagination = pagination_match.group(0)
    # 计算分页数
    links = re.findall(r'<a[^>]*href="([^"]*)"[^>]*>(\d+)</a>', pagination)
    if links:
        print(f"📋 分页链接:")
        for href, num in links[:10]:
            print(f"   第 {num} 页: ...{href[-80:]}")
    
    # 查找"下一页"
    next_match = re.search(r'<li[^>]*>\s*<a[^>]*href="([^"]*)"[^>]*>下一页|Next', pagination, re.IGNORECASE)
    if next_match:
        print(f"\n   下一页链接: {next_match.group(1)}")

# 查看表格内容
table_match = re.search(r'<table[^>]*>.*?</table>', html, re.DOTALL)
if table_match:
    table = table_match.group(0)
    rows = re.findall(r'<tr[^>]*>.*?</tr>', table, re.DOTALL)
    print(f"\n📊 表格有 {len(rows)} 行")
    
    # 显示前 5 行的内容
    print("\n表格内容（前 5 行）:")
    for i, row in enumerate(rows[:5], 1):
        cells = re.findall(r'<td[^>]*>([^<]*)</td>|<th[^>]*>([^<]*)</th>', row)
        cell_texts = [c[0] or c[1] for c in cells]
        cell_texts_cleaned = [c.strip() for c in cell_texts if c.strip()]
        if cell_texts_cleaned:
            print(f"  第 {i} 行: {' | '.join(cell_texts_cleaned[:5])}")

print("\n✅ 页面信息分析完成")
