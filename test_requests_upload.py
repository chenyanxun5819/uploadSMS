#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 requests 版本的成绩上传
"""

import sys
from pathlib import Path

# 添加项目目录到路径
sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "sms_app" / "core"))

from sms_handler_requests_v1 import SMSHandlerRequests


def test_upload():
    """测试成绩上传"""
    
    # 测试凭证
    username = "schhs334"
    password = "schhs334"
    
    # Excel 文件路径
    excel_path = str(Path(__file__).parent / "Upload.xlsx")
    
    print(">>> TEST SMS requests-based upload")
    print(f"    Username: {username}")
    print(f"    Excel: ...Upload.xlsx")
    print()
    
    # 创建处理器
    handler = SMSHandlerRequests(excel_path=excel_path)
    
    # 执行上传
    result = handler.upload_student_scores(
        username=username,
        password=password,
        year="2026",
        semester="1",
        date="2026-05-28",
        item_id=""  # 自动选择第一个项目
    )
    
    # 输出结果
    print()
    print(f"\n{'='*60}")
    print(f">>> UPLOAD RESULT")
    print(f"{'='*60}")
    print(f"Status: {'SUCCESS' if result['success'] else 'FAILED'}")
    print(f"Uploaded: {result['uploaded']} / {result['total']}")
    print(f"Message: {result['message']}")
    
    if result['errors']:
        print(f"Errors:")
        for error in result['errors'][:10]:  # 只显示前 10 个错误
            print(f"  - {error}")
    
    return result['success']


if __name__ == '__main__':
    success = test_upload()
    sys.exit(0 if success else 1)
