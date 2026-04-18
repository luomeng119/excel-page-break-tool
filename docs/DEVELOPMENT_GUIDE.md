# Excel分页符工具 - 开发指南

## 文档信息
- 版本：V2
- 创建日期：2026-04-17
- 用途：开发阶段技术文档

---

## 1. 开发环境准备

### 1.1 环境要求
- Python 3.10+
- Windows 10/11 (目标部署平台)
- 4GB+ RAM

### 1.2 初始化命令

```bash
# 创建虚拟环境
python -m venv venv

# 激活环境
venv\Scripts\activate  # Windows

# 安装依赖
pip install PyQt6==6.4.0 openpyxl==3.1.0 pyinstaller==6.0.0

# 验证安装
python -c "import PyQt6; import openpyxl; print('OK')"
```

### 1.3 项目结构初始化

```bash
mkdir excel_page_break_tool
cd excel_page_break_tool

mkdir ui core models utils resources build
cd resources
mkdir icons styles fonts

cd ..
touch requirements.txt
```

---

## 2. 核心模块开发顺序

### Phase 1: 基础框架（Day 1）

#### 2.1.1 配置文件 (config.py)
```python
"""全局配置"""

APP_NAME = "Excel分页符工具"
APP_VERSION = "2.0.0"

# 支持的文件格式
SUPPORTED_EXTENSIONS = ['.xlsx']

# 序号识别模式
SERIAL_PATTERNS = {
    '2位': '00',
    '3位': '000', 
    '4位': '0000',
    '自动': 'auto'
}

# 输出命名模板
OUTPUT_TEMPLATE = "{filename}_{date}_V2.xlsx"

# UI配置
WINDOW_MIN_SIZE = (1000, 700)
LEFT_PANEL_WIDTH = 350
```

#### 2.1.2 主入口 (main.py)
```python
"""程序入口"""
import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from ui.main_window import MainWindow
from config import APP_NAME

def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    
    # 启用高分屏支持
    app.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
```

#### 2.1.3 主窗口框架 (ui/main_window.py)
```python
"""主窗口 - 左右分栏布局"""
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, 
    QSplitter, QMessageBox
)
from PyQt6.QtCore import Qt
from config import WINDOW_MIN_SIZE, LEFT_PANEL_WIDTH
from ui.left_panel import LeftPanel
from ui.preview_panel import PreviewPanel

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setMinimumSize(*WINDOW_MIN_SIZE)
        self.setWindowTitle("Excel分页符工具 V2.0")
        
        self._setup_ui()
        self._connect_signals()
    
    def _setup_ui(self):
        """设置UI布局"""
        central = QWidget()
        self.setCentralWidget(central)
        
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # 分割器实现左右分栏
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 左侧控制面板
        self.left_panel = LeftPanel()
        self.left_panel.setMinimumWidth(LEFT_PANEL_WIDTH)
        self.left_panel.setMaximumWidth(LEFT_PANEL_WIDTH + 100)
        
        # 右侧预览面板
        self.preview_panel = PreviewPanel()
        
        splitter.addWidget(self.left_panel)
        splitter.addWidget(self.preview_panel)
        splitter.setSizes([LEFT_PANEL_WIDTH, 800])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        
        layout.addWidget(splitter)
    
    def _connect_signals(self):
        """连接信号槽"""
        self.left_panel.file_selected.connect(self.preview_panel.load_file)
        self.left_panel.process_started.connect(self._on_process_started)
        self.left_panel.progress_updated.connect(self._on_progress_updated)
    
    def _on_process_started(self, file_list):
        """开始处理"""
        pass
    
    def _on_progress_updated(self, current, total):
        """进度更新"""
        pass
```

---

### Phase 2: 左侧面板（Day 1-2）

#### 2.2.1 左侧面板 (ui/left_panel.py)
```python
"""左侧控制面板"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QGroupBox, 
    QRadioButton, QPushButton, QLineEdit,
    QProgressBar, QTextEdit, QCheckBox,
    QFileDialog, QLabel
)
from PyQt6.QtCore import pyqtSignal

class LeftPanel(QWidget):
    # 信号定义
    file_selected = pyqtSignal(list)  # 文件选择信号
    process_started = pyqtSignal(list)  # 开始处理信号
    progress_updated = pyqtSignal(int, int)  # 进度更新信号
    
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        
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
        layout = QVBoxLayout(group)
        
        # 模式选择
        self.single_mode = QRadioButton("单文件模式")
        self.batch_mode = QRadioButton("批量目录模式")
        self.batch_mode.setChecked(True)
        
        layout.addWidget(self.single_mode)
        layout.addWidget(self.batch_mode)
        
        # 路径选择
        path_layout = QHBoxLayout()
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("请选择文件或目录...")
        self.browse_btn = QPushButton("浏览...")
        self.browse_btn.clicked.connect(self._on_browse)
        
        path_layout.addWidget(self.path_input)
        path_layout.addWidget(self.browse_btn)
        layout.addLayout(path_layout)
        
        # 状态标签
        self.status_label = QLabel("状态: 就绪")
        layout.addWidget(self.status_label)
        
        parent_layout.addWidget(group)
    
    def _create_rule_group(self, parent_layout):
        """分页规则组"""
        group = QGroupBox("分页规则")
        layout = QVBoxLayout(group)
        
        # 自动补零选项
        self.auto_zero = QCheckBox("自动补零至3位")
        self.auto_zero.setChecked(True)
        layout.addWidget(self.auto_zero)
        
        # 红色标记选项
        self.red_mark = QCheckBox("标记背景色(红色)")
        self.red_mark.setChecked(True)
        layout.addWidget(self.red_mark)
        
        parent_layout.addWidget(group)
    
    def _create_progress_section(self, parent_layout):
        """进度显示"""
        group = QGroupBox("处理进度")
        layout = QVBoxLayout(group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("等待开始...")
        layout.addWidget(self.progress_label)
        
        parent_layout.addWidget(group)
    
    def _create_action_buttons(self, parent_layout):
        """操作按钮"""
        self.start_btn = QPushButton("开始处理")
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #1890ff;
                color: white;
                padding: 10px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
        """)
        self.start_btn.clicked.connect(self._on_start)
        parent_layout.addWidget(self.start_btn)
        
        self.preview_btn = QPushButton("打印预览")
        self.preview_btn.clicked.connect(self._on_preview)
        parent_layout.addWidget(self.preview_btn)
        
        self.open_dir_btn = QPushButton("打开输出目录")
        self.open_dir_btn.clicked.connect(self._on_open_dir)
        parent_layout.addWidget(self.open_dir_btn)
    
    def _create_log_section(self, parent_layout):
        """日志区域"""
        group = QGroupBox("日志")
        layout = QVBoxLayout(group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        layout.addWidget(self.log_text)
        
        parent_layout.addWidget(group)
    
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
                # TODO: 扫描目录获取文件列表
                self.status_label.setText("已选择目录")
    
    def _on_start(self):
        """开始处理"""
        if not self.selected_files:
            self.log("错误: 请先选择文件")
            return
        self.process_started.emit(self.selected_files)
    
    def _on_preview(self):
        """打印预览"""
        pass
    
    def _on_open_dir(self):
        """打开输出目录"""
        pass
    
    def log(self, message):
        """添加日志"""
        from datetime import datetime
        time_str = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{time_str}] {message}")
```

---

### Phase 3: 右侧预览面板（Day 2-3）

#### 2.3.1 预览面板 (ui/preview_panel.py)
```python
"""右侧预览面板"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QScrollArea,
    QFrame
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QPainter, QColor, QPen, QFont

class PreviewPanel(QWidget):
    def __init__(self):
        super().__init__()
        self.current_page = 1
        self.total_pages = 1
        self.zoom_level = 100
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # 工具栏
        toolbar = QHBoxLayout()
        
        self.zoom_out_btn = QPushButton("-")
        self.zoom_out_btn.setFixedWidth(30)
        self.zoom_out_btn.clicked.connect(self._zoom_out)
        
        self.zoom_label = QLabel("100%")
        self.zoom_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.zoom_label.setFixedWidth(50)
        
        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setFixedWidth(30)
        self.zoom_in_btn.clicked.connect(self._zoom_in)
        
        self.prev_btn = QPushButton("上一页")
        self.prev_btn.clicked.connect(self._prev_page)
        
        self.page_label = QLabel("第 1/1 页")
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.next_btn = QPushButton("下一页")
        self.next_btn.clicked.connect(self._next_page)
        
        toolbar.addWidget(self.zoom_out_btn)
        toolbar.addWidget(self.zoom_label)
        toolbar.addWidget(self.zoom_in_btn)
        toolbar.addStretch()
        toolbar.addWidget(self.prev_btn)
        toolbar.addWidget(self.page_label)
        toolbar.addWidget(self.next_btn)
        
        layout.addLayout(toolbar)
        
        # 预览画布
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        
        self.canvas = PreviewCanvas()
        scroll.setWidget(self.canvas)
        
        layout.addWidget(scroll)
    
    def load_file(self, file_path):
        """加载文件进行预览"""
        self.canvas.load_excel(file_path)
    
    def _zoom_in(self):
        self.zoom_level = min(200, self.zoom_level + 20)
        self.zoom_label.setText(f"{self.zoom_level}%")
        self.canvas.set_zoom(self.zoom_level)
    
    def _zoom_out(self):
        self.zoom_level = max(50, self.zoom_level - 20)
        self.zoom_label.setText(f"{self.zoom_level}%")
        self.canvas.set_zoom(self.zoom_level)
    
    def _prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self._update_page_label()
    
    def _next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self._update_page_label()
    
    def _update_page_label(self):
        self.page_label.setText(f"第 {self.current_page}/{self.total_pages} 页")


class PreviewCanvas(QWidget):
    """预览画布 - 自定义绘制"""
    
    def __init__(self):
        super().__init__()
        self.zoom = 1.0
        self.excel_data = None
        self.page_breaks = []
        self.serial_rows = []
        self.setMinimumSize(600, 800)
    
    def load_excel(self, file_path):
        """加载Excel数据"""
        # TODO: 使用openpyxl读取数据
        pass
    
    def set_zoom(self, percentage):
        """设置缩放比例"""
        self.zoom = percentage / 100.0
        self.update()
    
    def paintEvent(self, event):
        """绘制事件"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 背景
        painter.fillRect(self.rect(), QColor(240, 240, 240))
        
        # 绘制纸张区域
        paper_width = int(500 * self.zoom)
        paper_height = int(700 * self.zoom)
        paper_x = (self.width() - paper_width) // 2
        paper_y = 20
        
        painter.fillRect(paper_x, paper_y, paper_width, paper_height, Qt.GlobalColor.white)
        painter.setPen(QPen(QColor(200, 200, 200), 1))
        painter.drawRect(paper_x, paper_y, paper_width, paper_height)
        
        # TODO: 绘制Excel内容、分页线、红色标记
        
        painter.end()
```

---

### Phase 4: 核心处理模块（Day 3-4）

#### 2.4.1 序号检测器 (core/serial_number_detector.py)
```python
"""序号检测器 - 识别2/3/4位序号"""
from typing import List, NamedTuple
import re

class RowInfo(NamedTuple):
    row_index: int
    value: int
    format_code: str
    display_digits: int

class SerialNumberDetector:
    """检测Excel中的序号行"""
    
    # 支持的序号格式
    PATTERNS = [
        (r'^0+$', 2),    # 2位: 00
        (r'^0+$', 3),    # 3位: 000
        (r'^0+$', 4),    # 4位: 0000
    ]
    
    def detect(self, ws) -> List[RowInfo]:
        """
        检测工作表中的序号行
        
        Args:
            ws: openpyxl工作表对象
            
        Returns:
            List[RowInfo]: 检测到的序号行信息列表
        """
        results = []
        
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=1)
            
            if cell.value != 1:
                continue
            
            num_format = cell.number_format or ''
            
            # 检查是否为数字格式 (00, 000, 0000等)
            if self._is_serial_format(num_format):
                digits = len(num_format)
                results.append(RowInfo(
                    row_index=row_idx,
                    value=cell.value,
                    format_code=num_format,
                    display_digits=digits
                ))
        
        return results
    
    def _is_serial_format(self, num_format: str) -> bool:
        """检查是否为序号格式"""
        # 匹配纯0的格式字符串
        return bool(re.match(r'^0+$', num_format))
```

#### 2.4.2 Excel处理器 (core/excel_processor.py)
```python
"""Excel处理核心"""
from typing import List, Callable
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.pagebreak import Break
from datetime import datetime
import os

class ExcelProcessor:
    """Excel分页符处理器"""
    
    RED_FILL = PatternFill(
        start_color='FF0000',
        end_color='FF0000',
        fill_type='solid'
    )
    
    def __init__(self):
        self.progress_callback: Callable[[int, int], None] = None
    
    def set_progress_callback(self, callback: Callable[[int, int], None]):
        """设置进度回调函数"""
        self.progress_callback = callback
    
    def process_file(self, input_path: str, output_dir: str = None) -> str:
        """
        处理单个Excel文件
        
        Args:
            input_path: 输入文件路径
            output_dir: 输出目录（默认与输入文件同目录）
            
        Returns:
            str: 输出文件路径
        """
        # 加载工作簿
        wb = load_workbook(input_path)
        ws = wb.active
        
        # 检测序号行
        from core.serial_number_detector import SerialNumberDetector
        detector = SerialNumberDetector()
        serial_rows = detector.detect(ws)
        
        total_rows = len(serial_rows)
        
        # 处理每一行
        for idx, row_info in enumerate(serial_rows):
            # 添加分页符
            break_row = row_info.row_index - 1
            ws.row_breaks.append(Break(id=break_row))
            
            # 标记背景色
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_info.row_index, column=col)
                cell.fill = self.RED_FILL
            
            # 更新进度
            if self.progress_callback:
                self.progress_callback(idx + 1, total_rows)
        
        # 生成输出路径
        output_path = self._generate_output_path(input_path, output_dir)
        
        # 保存
        wb.save(output_path)
        
        return output_path
    
    def _generate_output_path(self, input_path: str, output_dir: str = None) -> str:
        """生成输出文件路径"""
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        date_str = datetime.now().strftime('%Y%m%d')
        file_name = f"{base_name}_{date_str}_V2.xlsx"
        
        if output_dir:
            return os.path.join(output_dir, file_name)
        else:
            dir_path = os.path.dirname(input_path)
            return os.path.join(dir_path, file_name)
```

---

### Phase 5: 多线程处理（Day 4）

#### 2.5.1 工作线程 (core/worker.py)
```python
"""后台工作线程"""
from PyQt6.QtCore import QThread, pyqtSignal
from typing import List
from core.excel_processor import ExcelProcessor

class ProcessWorker(QThread):
    """处理工作线程"""
    
    # 信号定义
    progress = pyqtSignal(int, int)  # 当前进度, 总数
    file_completed = pyqtSignal(str, str)  # 输入路径, 输出路径
    file_failed = pyqtSignal(str, str)  # 输入路径, 错误信息
    all_completed = pyqtSignal()
    
    def __init__(self, file_list: List[str]):
        super().__init__()
        self.file_list = file_list
        self.processor = ExcelProcessor()
        self.processor.set_progress_callback(self._on_progress)
        self._current_file_index = 0
    
    def _on_progress(self, current, total):
        """进度回调"""
        self.progress.emit(current, total)
    
    def run(self):
        """线程执行"""
        for file_path in self.file_list:
            try:
                output_path = self.processor.process_file(file_path)
                self.file_completed.emit(file_path, output_path)
            except Exception as e:
                self.file_failed.emit(file_path, str(e))
        
        self.all_completed.emit()
```

---

### Phase 6: 打包配置（Day 5）

#### 2.6.1 PyInstaller配置 (build/main.spec)
```python
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['../main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('../resources', 'resources'),
    ],
    hiddenimports=[
        'openpyxl',
        'PyQt6',
        'PyQt6.sip',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Excel分页符工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='../resources/icons/app.ico',
)
```

#### 2.6.2 打包命令
```bash
# 进入build目录
cd build

# 执行打包
pyinstaller main.spec --clean

# 输出位置
# dist/Excel分页符工具.exe
```

---

## 3. 测试规范

### 3.1 单元测试

#### 3.1.1 序号检测器测试 (tests/test_serial_detector.py)
```python
import unittest
from core.serial_number_detector import SerialNumberDetector
from openpyxl import Workbook

class TestSerialNumberDetector(unittest.TestCase):
    def setUp(self):
        self.detector = SerialNumberDetector()
        self.wb = Workbook()
        self.ws = self.wb.active
    
    def test_detect_2digit(self):
        """测试2位序号识别"""
        self.ws['A1'] = 1
        self.ws['A1'].number_format = '00'
        results = self.detector.detect(self.ws)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].display_digits, 2)
    
    def test_detect_3digit(self):
        """测试3位序号识别"""
        self.ws['A1'] = 1
        self.ws['A1'].number_format = '000'
        results = self.detector.detect(self.ws)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].display_digits, 3)
    
    def test_detect_4digit(self):
        """测试4位序号识别"""
        self.ws['A1'] = 1
        self.ws['A1'].number_format = '0000'
        results = self.detector.detect(self.ws)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].display_digits, 4)
    
    def test_ignore_normal_number(self):
        """测试普通数字不识别"""
        self.ws['A1'] = 1
        self.ws['A1'].number_format = 'General'
        results = self.detector.detect(self.ws)
        self.assertEqual(len(results), 0)
```

#### 3.1.2 Excel处理器测试 (tests/test_excel_processor.py)
```python
import unittest
import os
import tempfile
from core.excel_processor import ExcelProcessor
from openpyxl import load_workbook

class TestExcelProcessor(unittest.TestCase):
    def setUp(self):
        self.processor = ExcelProcessor()
        self.temp_dir = tempfile.mkdtemp()
    
    def test_process_adds_page_breaks(self):
        """测试是否正确添加分页符"""
        # 创建测试文件
        # ... 测试代码
        pass
    
    def test_process_marks_red_background(self):
        """测试是否正确标记红色背景"""
        # ... 测试代码
        pass
    
    def test_output_naming(self):
        """测试输出文件命名"""
        # ... 测试代码
        pass
```

### 3.2 集成测试

#### 3.2.1 UI集成测试
| 测试场景 | 测试步骤 | 验证点 |
|---------|---------|--------|
| 文件选择到处理 | 1.点击浏览 2.选择文件 3.点击开始 | 进度条更新，日志输出，结果文件生成 |
| 预览联动 | 选择文件后查看预览区 | 预览区正确显示Excel内容 |
| 批量处理 | 选择目录后开始处理 | 所有文件依次处理，进度正确 |

#### 3.2.2 多线程测试
```python
def test_concurrent_processing():
    """测试多线程处理不阻塞UI"""
    # 启动处理
    # 验证UI仍可响应
    # 验证进度更新
    pass
```

### 3.3 UI测试

#### 3.3.1 布局测试清单
- [ ] 窗口启动时默认尺寸正确 (1000x700)
- [ ] 左右分栏比例正确 (35%:65%)
- [ ] 分割器可拖动调整宽度
- [ ] 左侧面板最小宽度限制有效
- [ ] 预览区随窗口大小自适应
- [ ] 高分屏(4K)显示清晰

#### 3.3.2 交互测试清单
- [ ] 按钮点击有视觉反馈
- [ ] 进度条动画流畅
- [ ] 日志区域自动滚动
- [ ] 预览区缩放平滑
- [ ] 翻页按钮状态正确（首页禁用上一页，末页禁用下一页）

### 3.4 端到端测试

#### 3.4.1 完整流程测试
```
测试名称: 单文件完整流程
前置条件: 准备好包含001序号的测试Excel文件

步骤:
1. 启动程序
2. 选择"单文件模式"
3. 点击"浏览"选择测试文件
4. 点击"开始处理"
5. 等待处理完成
6. 查看预览区
7. 打开输出文件验证

验证点:
- 输出文件生成成功
- 分页符添加正确
- 红色背景标记正确
- 预览区显示正确
```

#### 3.4.2 批量流程测试
```
测试名称: 批量目录处理
前置条件: 目录中包含5个不同的Excel文件

步骤:
1. 选择"批量目录模式"
2. 选择测试目录
3. 点击"开始处理"
4. 观察进度条和日志
5. 等待全部完成

验证点:
- 所有文件都被处理
- 每个文件生成对应输出
- 进度条从0%到100%
- 日志记录每个文件处理状态
```

### 3.5 性能测试

#### 3.5.1 性能基准
| 测试项 | 测试数据 | 目标值 |
|-------|---------|--------|
| 启动时间 | 冷启动 | < 3秒 |
| 文件加载 | 1000行Excel | < 1秒 |
| 处理速度 | 1000行 | < 5秒 |
| 内存占用 | 处理10000行 | < 500MB |
| 预览渲染 | 100页 | < 2秒 |

#### 3.5.2 压力测试
```python
def test_large_file_processing():
    """测试大文件处理性能"""
    # 创建10000行测试数据
    # 测量处理时间和内存占用
    # 验证无内存泄漏
    pass

def test_batch_processing():
    """测试批量处理性能"""
    # 准备50个文件
    # 测量总处理时间
    # 验证进度更新正常
    pass
```

### 3.6 兼容性测试

| 环境 | 版本 | 测试内容 |
|-----|------|---------|
| Windows 10 | 21H2, 22H2 | 安装、运行、功能完整 |
| Windows 11 | 23H2 | 安装、运行、功能完整 |
| Excel | 2016, 2019, 365 | 文件读取、分页符添加 |
| WPS | 最新版 | 文件读取、分页符添加 |
| Python | 3.10, 3.11, 3.12 | 源码运行正常 |

### 3.7 开发检查清单

#### 功能检查
- [ ] 单文件选择正常
- [ ] 批量目录扫描正常
- [ ] 2位序号(01)识别正确
- [ ] 3位序号(001)识别正确
- [ ] 4位序号(0001)识别正确
- [ ] 分页符添加正确
- [ ] 红色背景标记正确
- [ ] 预览渲染正确
- [ ] 缩放功能正常
- [ ] 翻页功能正常

#### 布局检查
- [ ] 左右分栏比例正确
- [ ] 分割器拖动流畅
- [ ] 窗口大小自适应
- [ ] 高分屏显示清晰

#### 流程检查
- [ ] 单文件流程完整
- [ ] 批量流程完整
- [ ] 错误处理正确
- [ ] 进度更新及时

#### 异常检查
- [ ] 文件被占用提示正确
- [ ] 非Excel文件过滤正确
- [ ] 大文件处理不卡顿
- [ ] 内存占用合理

#### 打包检查
- [ ] 单文件exe生成成功
- [ ] 干净环境运行正常
- [ ] 图标显示正确

---

## 4. 参考资源

### 官方文档
- PyQt6: https://doc.qt.io/qtforpython-6/
- openpyxl: https://openpyxl.readthedocs.io/
- PyInstaller: https://pyinstaller.org/

### 关键技术点
1. **QSplitter**: 实现左右分栏可拖拽调整
2. **QPainter**: 自定义绘制预览内容
3. **QThread**: 后台处理不阻塞UI
4. **pyqtSignal**: 线程间通信
