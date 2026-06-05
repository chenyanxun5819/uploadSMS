#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS 学生成绩自动上传系统 - 主应用入口
"""

import sys
import os
from pathlib import Path
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QTimer

from ui.main_window import MainWindow
from core.startup_checker import StartupChecker


def main():
    app = QApplication(sys.argv)
    
    # 设置应用样式 - 兼容 PyInstaller 打包环境
    if getattr(sys, 'frozen', False):
        # PyInstaller 环境
        app_dir = Path(sys._MEIPASS) / 'ui'
    else:
        # 开发环境
        app_dir = Path(__file__).parent / 'ui'
    
    styles_path = app_dir / 'styles.qss'
    if styles_path.exists():
        with open(styles_path, 'r', encoding='utf-8') as f:
            app.setStyleSheet(f.read())
    else:
        print(f" 警告: 样式文件未找到: {styles_path}")
    
    # 创建主窗口并最大化显示（保留窗口控制）
    window = MainWindow()
    window.showMaximized()
    
    # 显示启动日志和日志文件位置
    log_dir = Path.home() / ".sms_app" / "logs"
    today = __import__('datetime').datetime.now().strftime("%Y-%m-%d")
    log_file = log_dir / f"sms_app_{today}.log"
    window.console.log_success("SMS 学生成绩自动上传系统已启动")
    window.console.log_info(f"📂 日志保存位置: {log_file}", "#8abaff")
    
    # 延迟执行启动检查（避免阻塞UI）
    def perform_startup_check():
        checker = StartupChecker()
        
        def log_callback(message):
            """将检查日志输出到console"""
            # 根据消息类型选择日志级别
            if "✅" in message or "已启动" in message:
                window.console.log_success(message)
            elif "❌" in message or "失败" in message:
                window.console.log_error(message)
            elif "⚠️" in message or "警告" in message or "未保存" in message:
                window.console.log_warning(message)
            elif "=" in message:
                window.console.log_info(message, "#6a9fb5")
            else:
                window.console.log_info(message, "#8abaff")
        
        # 使用增量检查（更快，只检查差异部分）
        result = checker.check_and_update_incremental(log_callback=log_callback)
        
        # 输出最终结果摘要
        if result['checked']:
            if result['matched']:
                window.console.log_success(f"✅ 数据检查完成 - {result['message']}")
            else:
                if result['updated']:
                    window.console.log_success(f"✅ 数据已更新 - {result['message']}")
                else:
                    window.console.log_error(f"❌ 数据更新失败 - {result['message']}")
        else:
            window.console.log_warning(f"⚠️  数据检查跳过 - {result['message']}")
    
    # 使用定时器在UI准备好后执行检查
    timer = QTimer()
    timer.setSingleShot(True)
    timer.timeout.connect(perform_startup_check)
    timer.start(500)  # 延迟500ms执行
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
