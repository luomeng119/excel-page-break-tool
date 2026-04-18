"""左侧控制面板"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox,
    QRadioButton, QPushButton, QLineEdit,
    QProgressBar, QTextEdit, QCheckBox,
    QFileDialog, QLabel, QHBoxLayout,
    QMessageBox, QComboBox, QColorDialog
)
from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtGui import QColor
import os

from config import COLOR_OPTIONS, DEFAULT_COLOR, APP_NAME, APP_VERSION


class LeftPanel(QWidget):
    """左侧控制面板"""

    # 信号定义
    file_selected = pyqtSignal(list)  # 文件选择信号
    process_started = pyqtSignal(list, str, str)  # 开始处理信号（文件列表，输出目录，标记颜色）
    process_stopped = pyqtSignal()  # 停止处理信号
    preview_requested = pyqtSignal()  # 预览请求信号
    open_output_dir_requested = pyqtSignal()  # 打开输出目录信号
    color_changed = pyqtSignal(str)  # 颜色变化信号

    def __init__(self):
        super().__init__()
        self.selected_files = []
        from config import OUTPUT_DIR
        self.output_dir = OUTPUT_DIR
        self.is_processing = False
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(15, 15, 15, 15)

        # 产品标题
        title_label = QLabel(f"{APP_NAME}")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #1890ff;
                padding: 10px 0;
            }
        """)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        # 版本号
        version_label = QLabel(f"{APP_VERSION}")
        version_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #999;
                padding-bottom: 5px;
            }
        """)
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(version_label)

        # 1. 输入设置组
        self._create_input_group(layout)
        
        # 2. 分页规则组
        self._create_rule_group(layout)
        
        # 3. 进度显示
        self._create_progress_section(layout)
        
        # 4. 操作按钮
        self._create_action_buttons(layout)
        
        # 5. 日志区域
        self._create_log_section(layout)
        
        layout.addStretch()
    
    def _create_input_group(self, parent_layout):
        """输入设置组"""
        group = QGroupBox("输入设置")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QVBoxLayout(group)
        layout.setSpacing(10)
        
        # 模式选择
        mode_layout = QHBoxLayout()
        self.single_mode = QRadioButton("单文件模式")
        self.batch_mode = QRadioButton("批量目录模式")
        self.batch_mode.setChecked(True)
        self.single_mode.toggled.connect(self._on_mode_changed)
        
        mode_layout.addWidget(self.single_mode)
        mode_layout.addWidget(self.batch_mode)
        mode_layout.addStretch()
        layout.addLayout(mode_layout)
        
        # 路径选择
        path_layout = QHBoxLayout()
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("请选择文件或目录...")
        self.path_input.setReadOnly(True)
        self.path_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background-color: #f5f5f5;
            }
        """)
        
        self.browse_btn = QPushButton("浏览...")
        self.browse_btn.setFixedWidth(70)
        self.browse_btn.setStyleSheet("""
            QPushButton {
                padding: 6px 12px;
                background-color: #1890ff;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
        """)
        self.browse_btn.clicked.connect(self._on_browse)
        
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.browse_btn)
        layout.addLayout(path_layout)
        
        # 状态标签
        self.status_label = QLabel("状态: 就绪")
        self.status_label.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(self.status_label)
        
        parent_layout.addWidget(group)
    
    def _create_rule_group(self, parent_layout):
        """分页规则组"""
        group = QGroupBox("分页规则")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QVBoxLayout(group)
        
        # 自动补零选项
        self.auto_zero = QCheckBox("自动补零至3位")
        self.auto_zero.setChecked(True)
        self.auto_zero.setToolTip("将2位序号(01)自动补零显示为3位(001)")
        layout.addWidget(self.auto_zero)

        # 保留背景色选项
        self.keep_colors = QCheckBox("输出保留背景色")
        self.keep_colors.setChecked(True)
        self.keep_colors.setToolTip("勾选: 输出文件带颜色标记(校对用)\n取消: 输出仅含分页符(正式打印)")
        layout.addWidget(self.keep_colors)
        
        # 颜色选择
        color_layout = QHBoxLayout()
        color_label = QLabel("标记颜色:")
        color_layout.addWidget(color_label)
        
        self.color_combo = QComboBox()
        for color_name in COLOR_OPTIONS.keys():
            self.color_combo.addItem(color_name)
        self.color_combo.setCurrentText('红色')
        self.color_combo.setStyleSheet("""
            QComboBox {
                padding: 4px;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background-color: white;
            }
        """)
        color_layout.addWidget(self.color_combo)
        
        # 颜色预览方块
        self.color_preview = QLabel("  ")
        self.color_preview.setFixedSize(24, 24)
        self.color_preview.setStyleSheet(f"background-color: #{DEFAULT_COLOR}; border: 1px solid #ccc; border-radius: 4px;")
        color_layout.addWidget(self.color_preview)
        
        # 自定义颜色按钮
        self.custom_color_btn = QPushButton("自定义")
        self.custom_color_btn.setFixedWidth(60)
        self.custom_color_btn.setStyleSheet("""
            QPushButton {
                padding: 4px 8px;
                background-color: #f5f5f5;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
            }
        """)
        self.custom_color_btn.clicked.connect(self._on_custom_color)
        color_layout.addWidget(self.custom_color_btn)
        
        color_layout.addStretch()
        layout.addLayout(color_layout)
        
        # 连接颜色选择变化
        self.color_combo.currentTextChanged.connect(self._on_color_changed)
        
        parent_layout.addWidget(group)
    
    def _create_progress_section(self, parent_layout):
        """进度显示"""
        group = QGroupBox("处理进度")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QVBoxLayout(group)
        
        # 文件进度
        self.file_progress_label = QLabel("文件: 0/0")
        self.file_progress_label.setStyleSheet("font-size: 12px; color: #666;")
        layout.addWidget(self.file_progress_label)
        
        # 总体进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                text-align: center;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #1890ff;
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
        # 当前进度标签
        self.progress_label = QLabel("等待开始...")
        self.progress_label.setStyleSheet("font-size: 12px; color: #999;")
        layout.addWidget(self.progress_label)
        
        parent_layout.addWidget(group)
    
    def _create_action_buttons(self, parent_layout):
        """操作按钮"""
        btn_layout = QVBoxLayout()
        btn_layout.setSpacing(10)
        
        self.start_btn = QPushButton("开始处理")
        self.start_btn.setMinimumHeight(40)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #52c41a;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #73d13d;
            }
            QPushButton:pressed {
                background-color: #389e0d;
            }
            QPushButton:disabled {
                background-color: #d9d9d9;
                color: #999;
            }
        """)
        self.start_btn.clicked.connect(self._on_start)
        
        self.preview_btn = QPushButton("刷新预览")
        self.preview_btn.setMinimumHeight(35)
        self.preview_btn.setStyleSheet("""
            QPushButton {
                background-color: #1890ff;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
            QPushButton:disabled {
                background-color: #d9d9d9;
                color: #999;
            }
        """)
        self.preview_btn.clicked.connect(self._on_preview)
        
        self.open_dir_btn = QPushButton("打开输出目录")
        self.open_dir_btn.setMinimumHeight(35)
        self.open_dir_btn.setStyleSheet("""
            QPushButton {
                background-color: #faad14;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #ffc53d;
            }
        """)
        self.open_dir_btn.clicked.connect(self._on_open_dir)
        
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.preview_btn)
        btn_layout.addWidget(self.open_dir_btn)
        
        parent_layout.addLayout(btn_layout)
    
    def _create_log_section(self, parent_layout):
        """日志区域"""
        group = QGroupBox("处理日志")
        group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        layout = QVBoxLayout(group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                background-color: #fafafa;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 11px;
            }
        """)
        layout.addWidget(self.log_text)
        
        # 清空日志按钮
        clear_btn = QPushButton("清空日志")
        clear_btn.setFixedHeight(25)
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                border: 1px solid #d9d9d9;
                border-radius: 4px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
            }
        """)
        clear_btn.clicked.connect(self.clear_log)
        layout.addWidget(clear_btn)
        
        parent_layout.addWidget(group)
    
    def _on_mode_changed(self):
        """模式切换"""
        self.selected_files = []
        self.path_input.clear()
        self.status_label.setText("状态: 就绪")
    
    def _on_browse(self):
        """浏览按钮点击"""
        if self.single_mode.isChecked():
            files, _ = QFileDialog.getOpenFileNames(
                self, "选择Excel文件", "", "Excel文件 (*.xlsx)"
            )
            if files:
                self.selected_files = files
                self.path_input.setText(files[0])
                self.status_label.setText(f"已选择 {len(files)} 个文件")
                self.file_selected.emit(files)
        else:
            directory = QFileDialog.getExistingDirectory(
                self, "选择目录"
            )
            if directory:
                self.path_input.setText(directory)
                # 扫描目录获取文件列表
                self.selected_files = self._scan_directory(directory)
                self.status_label.setText(f"已选择目录，找到 {len(self.selected_files)} 个Excel文件")
                if self.selected_files:
                    self.file_selected.emit(self.selected_files)
    
    def _scan_directory(self, directory: str) -> list:
        """扫描目录获取Excel文件"""
        files = []
        for root, dirs, filenames in os.walk(directory):
            for filename in filenames:
                if filename.endswith('.xlsx') and not filename.startswith('~$'):
                    files.append(os.path.join(root, filename))
        return files
    
    def _on_start(self):
        """开始处理"""
        if self.is_processing:
            # 停止处理
            self.is_processing = False
            self.start_btn.setText("开始处理")
            self.start_btn.setStyleSheet("""
                QPushButton {
                    background-color: #52c41a;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    font-size: 14px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #73d13d;
                }
            """)
            self.process_stopped.emit()
            return
        
        if not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择文件或目录")
            return
        
        self.is_processing = True
        self.start_btn.setText("停止处理")
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff4d4f;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #ff7875;
            }
        """)
        
        # 获取选中的颜色
        selected_color = self.get_selected_color()
        self.process_started.emit(self.selected_files, self.output_dir, selected_color)
    
    def _on_preview(self):
        """打印预览"""
        self.preview_requested.emit()
    
    def _on_open_dir(self):
        """打开输出目录"""
        self.open_output_dir_requested.emit()
    
    def _on_color_changed(self, color_name):
        """颜色选择变化"""
        if color_name in COLOR_OPTIONS:
            color_hex = COLOR_OPTIONS[color_name]
            self.color_preview.setStyleSheet(f"background-color: #{color_hex}; border: 1px solid #ccc; border-radius: 4px;")
            # 发送颜色变化信号，通知预览面板更新
            self.color_changed.emit(color_hex)
    
    def _on_custom_color(self):
        """自定义颜色"""
        color = QColorDialog.getColor()
        if color.isValid():
            color_hex = color.name().lstrip('#').upper()
            self.color_preview.setStyleSheet(f"background-color: #{color_hex}; border: 1px solid #ccc; border-radius: 4px;")
            # 添加到自定义颜色选项
            self.color_combo.addItem(f"自定义({color_hex})")
            self.color_combo.setCurrentText(f"自定义({color_hex})")
            # 保存到COLOR_OPTIONS
            COLOR_OPTIONS[f"自定义({color_hex})"] = color_hex
    
    def get_selected_color(self) -> str:
        """获取选中的颜色值"""
        color_name = self.color_combo.currentText()
        if color_name in COLOR_OPTIONS:
            return COLOR_OPTIONS[color_name]
        # 处理自定义颜色
        if "自定义(" in color_name:
            return color_name.split("(")[1].rstrip(")")
        return DEFAULT_COLOR
    
    def log(self, message: str):
        """添加日志"""
        from datetime import datetime
        time_str = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{time_str}] {message}")
        # 滚动到底部
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def update_progress(self, current: int, total: int):
        """更新进度"""
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress_bar.setValue(percentage)
            self.progress_label.setText(f"处理中: {current}/{total} ({percentage}%)")
    
    def update_file_progress(self, current: int, total: int):
        """更新文件进度"""
        self.file_progress_label.setText(f"文件: {current}/{total}")
    
    def reset_progress(self):
        """重置进度"""
        self.progress_bar.setValue(0)
        self.progress_label.setText("等待开始...")
        self.file_progress_label.setText("文件: 0/0")
        self.is_processing = False
        self.start_btn.setText("开始处理")
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #52c41a;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #73d13d;
            }
        """)
    
    def set_output_dir(self, output_dir: str):
        """设置输出目录"""
        self.output_dir = output_dir
