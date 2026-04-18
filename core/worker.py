"""后台工作线程"""
from PyQt6.QtCore import QThread, pyqtSignal
from typing import List
from core.excel_processor import ExcelProcessor


class ProcessWorker(QThread):
    """处理工作线程"""

    # 信号定义
    progress = pyqtSignal(int, int)  # 当前进度, 总数
    file_progress = pyqtSignal(int, int)  # 当前文件索引, 总文件数
    file_completed = pyqtSignal(str, str)  # 输入路径, 输出路径
    file_failed = pyqtSignal(str, str)  # 输入路径, 错误信息
    log_message = pyqtSignal(str)  # 日志消息
    all_completed = pyqtSignal()

    def __init__(self, file_list: List[str], output_dir: str = None, mark_color: str = 'FF0000', preview_mode: bool = False):
        super().__init__()
        self.file_list = file_list
        self.output_dir = output_dir
        self.processor = ExcelProcessor(mark_color=mark_color)
        self.processor.set_mark_color(mark_color)
        self.processor.set_preview_mode(preview_mode)
        self.processor.set_progress_callback(self._on_progress)
        self.processor.set_log_callback(self._on_log)
        self._is_running = True
    
    def _on_progress(self, current, total):
        """进度回调"""
        self.progress.emit(current, total)
    
    def _on_log(self, message):
        """日志回调"""
        self.log_message.emit(message)
    
    def stop(self):
        """停止处理"""
        self._is_running = False
        self.wait(1000)
    
    def run(self):
        """线程执行"""
        total_files = len(self.file_list)
        
        for file_idx, file_path in enumerate(self.file_list):
            if not self._is_running:
                self._on_log("处理已取消")
                break
            
            self.file_progress.emit(file_idx + 1, total_files)
            
            try:
                output_path = self.processor.process_file(file_path, self.output_dir)
                if output_path:
                    self.file_completed.emit(file_path, output_path)
                else:
                    self.file_failed.emit(file_path, "未检测到序号行")
            except Exception as e:
                self.file_failed.emit(file_path, str(e))
        
        self.all_completed.emit()
