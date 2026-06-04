#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试改进后的 sms_handler 学生上传功能
"""

import sys
from pathlib import Path

# Add sms_app to path
sys.path.insert(0, str(Path(__file__).parent / 'sms_app'))

from core.sms_handler import SMSHandler

handler = SMSHandler()

print("=" * 100)
print("测试改进的学生上传功能")
print("=" * 100)

# 测试上传
result = handler.upload_student_scores(
    username='schhs334',
    password='schhs334',
    date='2026-02-06',
    activity_code='ACA CMO207'
)

print("\n" + "=" * 100)
print("上传结果")
print("=" * 100)
print(f"状态: {'✓ 成功' if result['success'] else '✗ 失败'}")
print(f"消息: {result.get('message', '')}")
print(f"上传数: {result.get('total', 0)} / 3")
if result.get('errors'):
    print(f"错误:")
    for error in result['errors']:
        print(f"  - {error}")
