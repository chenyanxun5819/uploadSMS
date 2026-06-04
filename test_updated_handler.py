#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test updated SMSHandler with requests implementation
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "sms_app"))

from core.sms_handler import SMSHandler

def test_upload():
    username = "schhs334"
    password = "schhs334"
    
    print(">>> Testing updated SMSHandler (requests version)")
    print(f"    Username: {username}")
    print()
    
    handler = SMSHandler()
    result = handler.upload_student_scores(username=username, password=password)
    
    print()
    print(f"\n{'='*60}")
    print(f">>> RESULT")
    print(f"{'='*60}")
    print(f"Success: {result['success']}")
    print(f"Uploaded: {result['uploaded']} / {result['total']}")
    print(f"Failed: {result['failed']}")
    print(f"Message: {result['message']}")
    
    if result['errors'] and result['uploaded'] > 0:
        print(f"\nNot found (first 10):")
        for error in result['errors'][:10]:
            print(f"  - {error}")
    
    return result['success']

if __name__ == '__main__':
    success = test_upload()
    sys.exit(0 if success else 1)
