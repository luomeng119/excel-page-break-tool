"""程序入口"""
import sys
import os

# 添加项目目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from ui.main_window import MainWindow
from config import APP_NAME


def main():
    # 启用高分屏支持（必须在创建QApplication之前）
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion("2.0.0")
    
    # 设置全局字体
    font = QFont("Arial", 10)
    app.setFont(font)
    
    # 设置应用样式
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
