"""主窗口 - 左右分栏布局"""
import os
import sys

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout,
    QSplitter, QMessageBox, QFileDialog
)
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices

from config import WINDOW_MIN_SIZE, LEFT_PANEL_WIDTH, LEFT_PANEL_MAX_WIDTH, APP_NAME, APP_VERSION
from ui.left_panel import LeftPanel
from ui.preview_panel import PreviewPanel
from core.worker import ProcessWorker


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.output_dir = None
        self.processed_files = []
        self.last_processed_file = None

        self.setMinimumSize(*WINDOW_MIN_SIZE)
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.setStyleSheet("QMainWindow { background-color: #f5f5f5; }")

        self._setup_ui()
        self._connect_signals()

        self.left_panel.log("程序启动完成")
        self.left_panel.log("请选择文件或目录开始处理")

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setStyleSheet("""
            QSplitter::handle { background-color: #d9d9d9; }
            QSplitter::handle:horizontal { width: 4px; }
            QSplitter::handle:horizontal:hover { background-color: #1890ff; }
        """)

        self.left_panel = LeftPanel()
        self.left_panel.setMinimumWidth(LEFT_PANEL_WIDTH)
        self.left_panel.setMaximumWidth(LEFT_PANEL_MAX_WIDTH)
        self.left_panel.setStyleSheet("background-color: #ffffff;")

        self.preview_panel = PreviewPanel()
        self.preview_panel.setStyleSheet("background-color: #f5f5f5;")

        splitter.addWidget(self.left_panel)
        splitter.addWidget(self.preview_panel)
        splitter.setSizes([LEFT_PANEL_WIDTH, 800])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

        layout.addWidget(splitter)

    def _connect_signals(self):
        self.left_panel.file_selected.connect(self._on_file_selected)
        self.left_panel.process_started.connect(self._on_process_started)
        self.left_panel.process_stopped.connect(self._on_process_stopped)
        self.left_panel.preview_requested.connect(self._on_preview_requested)
        self.left_panel.open_output_dir_requested.connect(self._on_open_output_dir)
        self.left_panel.color_changed.connect(self._on_color_changed)

    def _on_file_selected(self, file_list):
        if file_list:
            self.left_panel.log(f"已选择 {len(file_list)} 个文件")
            self.preview_panel.load_file(file_list[0])
            # 选择新文件时清除处理记录
            self.last_processed_file = None

    def _on_process_started(self, file_list, output_dir, mark_color):
        if not file_list:
            return

        self.processed_files = []
        self.output_dir = output_dir
        self.last_processed_file = None

        self.left_panel.log(f"开始处理 {len(file_list)} 个文件...")
        self.left_panel.log(f"标记颜色: #{mark_color}")
        self.left_panel.reset_progress()

        # 获取"保留背景色"选项
        keep_colors = self.left_panel.keep_colors.isChecked()

        # 创建工作线程 - 根据用户选择决定是否标记颜色
        self.worker = ProcessWorker(file_list, output_dir, mark_color=mark_color, preview_mode=keep_colors)
        self.worker.progress.connect(self.left_panel.update_progress)
        self.worker.file_progress.connect(self.left_panel.update_file_progress)
        self.worker.file_completed.connect(self._on_file_completed)
        self.worker.file_failed.connect(self._on_file_failed)
        self.worker.log_message.connect(self.left_panel.log)
        self.worker.all_completed.connect(self._on_all_completed)

        self.worker.start()

    def _on_process_stopped(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.left_panel.log("处理已停止")

    def _on_file_completed(self, input_path, output_path):
        self.processed_files.append(output_path)
        self.last_processed_file = output_path

    def _on_file_failed(self, input_path, error_msg):
        self.left_panel.log(f"处理失败: {os.path.basename(input_path)} - {error_msg}")

    def _on_all_completed(self):
        self.left_panel.reset_progress()
        self.left_panel.log(f"处理完成，共生成 {len(self.processed_files)} 个文件")

        if self.last_processed_file:
            # 自动加载处理后的文件预览
            self.left_panel.log(f"正在加载预览: {os.path.basename(self.last_processed_file)}")
            self.preview_panel.load_processed_file(self.last_processed_file)

            QMessageBox.information(
                self,
                "处理完成",
                f"成功处理 {len(self.processed_files)} 个文件！\n"
                f"输出目录: {os.path.dirname(self.processed_files[0])}"
            )

    def _on_preview_requested(self):
        """刷新预览 - 优先显示处理后的文件"""
        if self.last_processed_file:
            self.preview_panel.load_processed_file(self.last_processed_file)
            self.left_panel.log("预览已刷新（处理后文件）")
        elif self.left_panel.selected_files:
            self.preview_panel.load_file(self.left_panel.selected_files[0])
            self.left_panel.log("预览已刷新（原始文件）")
        else:
            QMessageBox.warning(self, "警告", "请先选择文件")

    def _on_open_output_dir(self):
        if self.processed_files:
            output_dir = os.path.dirname(self.processed_files[0])
        elif self.left_panel.selected_files:
            output_dir = os.path.dirname(self.left_panel.selected_files[0])
        else:
            QMessageBox.warning(self, "警告", "没有可用的输出目录")
            return

        if os.path.exists(output_dir):
            QDesktopServices.openUrl(QUrl.fromLocalFile(output_dir))
        else:
            QMessageBox.warning(self, "警告", "输出目录不存在")

    def _on_color_changed(self, color_hex):
        """颜色变化 - 仅更新预览画布颜色，不重新加载文件"""
        self.preview_panel.set_highlight_color(color_hex)
        self.left_panel.log(f"预览颜色已更新: #{color_hex}")

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "确认退出", "正在处理中，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.worker.stop()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()
