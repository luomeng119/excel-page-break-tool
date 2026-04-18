"""全局配置"""
import os

APP_NAME = "Excel分页符工具"
APP_VERSION = "V1.0"

# 支持的文件格式
SUPPORTED_EXTENSIONS = ['.xlsx']

# 序号识别模式
SERIAL_PATTERNS = {
    '2位': '00',
    '3位': '000',
    '4位': '0000',
    '自动': 'auto'
}

# 输出目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'test')

# 确保输出目录存在
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 输出命名模板（PRD V1规范）
# 格式: {原文件名}_{YYYYMMDD}_V{版本号}.xlsx
# APP_VERSION = "V1.0" -> 输出文件名: 卷内目录_20260417_VV1.0.xlsx
# 为避免双V，更简洁的方案: 版本号不带V前缀
# 输出: 卷内目录_20260417_V1.0.xlsx
OUTPUT_TEMPLATE = "{filename}_{date}_V{version}.xlsx"

# UI配置
WINDOW_MIN_SIZE = (1000, 700)
LEFT_PANEL_WIDTH = 350
LEFT_PANEL_MAX_WIDTH = 450

# 七彩颜色配置
COLOR_OPTIONS = {
    '红色': 'FF0000',
    '橙色': 'FF7F00',
    '黄色': 'FFFF00',
    '绿色': '00FF00',
    '青色': '00FFFF',
    '蓝色': '0000FF',
    '紫色': '8B00FF',
}

# 默认颜色
DEFAULT_COLOR = 'FF0000'

# 预览背景配置
PREVIEW_BACKGROUND = (240, 240, 240)
PAPER_BACKGROUND = (255, 255, 255)
PAGE_BREAK_LINE = (255, 0, 0)
