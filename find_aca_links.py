#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
查看 ACA CMO5 项目的链接
"""

import re

with open('search_result_page1.html', 'r', encoding='utf-8') as f:
    html = f.read()

# 查找包含 ACA CMO5 的行及其相关链接
print("寻找 ACA CMO5 项目的链接...")
print()

matches = list(re.finditer(r'<tr[^>]*>.*?ACA CMO5.*?</tr>', html, re.DOTALL))

if matches:
    for i, match in enumerate(matches, 1):
        row = match.group(0)
        
        # 提取所有 href 链接
        hrefs = re.findall(r'href="([^"]*)"', row)
        
        print(f"第 {i} 个 ACA CMO5 记录:")
        print(f"  链接数: {len(hrefs)}")
        for j, href in enumerate(hrefs, 1):
            print(f"    {j}. {href}")
        
        # 提取单元格文本
        cells = re.findall(r'<td[^>]*>([^<]+)</td>', row)
        print(f"  单元格内容 ({len(cells)} 个):")
        for j, cell in enumerate(cells, 1):
            cell_clean = cell.strip()[:80]
            print(f"    {j}. {cell_clean}")
        
        print()
else:
    print("未找到 ACA CMO5 项目")

# 也检查所有包含 'studentPerformance' 的链接
print("\n" + "="*80)
print("检查是否存在指向成绩录入页面的链接...")
print("="*80 + "\n")

perf_links = re.findall(r'href="(/sms/index\.php\?[^"]*studentPerformance[^"]*)"', html)
if perf_links:
    print(f"找到 {len(perf_links)} 个 studentPerformance 链接:")
    for link in set(perf_links)[:5]:
        print(f"  {link}")
else:
    print("未找到 studentPerformance 链接")
