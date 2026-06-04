#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试修改后的 SMS 上传功能
"""

import sys
from pathlib import Path

# Add sms_app to path
sms_app_path = Path(__file__).parent / "sms_app"
sys.path.insert(0, str(sms_app_path))

from core.sms_handler import SMSHandler

def test_upload():
    """测试上传功能"""
    
    # 用户输入
    print("=" * 60)
    print("测试 SMS 学生成绩上传")
    print("=" * 60)
    
    username = input("\n请输入用户名: ").strip()
    password = input("请输入密码: ").strip()
    
    if not username or not password:
        print("ERROR: 用户名和密码不能为空")
        return
    
    # 初始化处理器
    handler = SMSHandler(headless=False)
    
    # 执行上传
    print("\n开始上传...")
    print("-" * 60)
    
    result = handler.upload_student_scores(
        username=username,
        password=password,
        date='2026-02-06',
        activity_code='ACA CMO207'
    )
    
    print("-" * 60)
    print("\n上传结果:")
    print(f"  成功: {result['success']}")
    print(f"  已上传: {result['uploaded']}")
    print(f"  失败: {result['failed']}")
    print(f"  总数: {result['total']}")
    print(f"  消息: {result['message']}")
    
    if result['errors']:
        print(f"  错误详情:")
        for error in result['errors']:
            print(f"    - {error}")

if __name__ == '__main__':
    try:
        test_upload()
    except KeyboardInterrupt:
        print("\n\n用户中断")
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
