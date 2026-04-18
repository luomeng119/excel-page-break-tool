# Excel分页符工具 V1.0

面向档案管理人员的 Excel 分页符自动化处理桌面工具。自动识别卷内目录中的序号行，插入分页符，标记颜色，支持打印预览。

## 产品功能

### 核心功能
- **智能序号识别** - 自动识别 2位(01)、3位(001)、4位(0001) 格式的序号行
- **自动分页符** - 在每个新卷序号行上方插入 Excel 分页符
- **颜色标记** - 序号行背景色高亮（7色可选+自定义），输出可选保留或去除
- **自动备份** - 处理前自动备份原始文件

### 界面功能
- **打印预览** - 右侧实时预览分页效果，支持缩放和翻页
- **批量处理** - 支持单文件和目录批量模式
- **颜色实时切换** - 左侧选色，右侧预览即时变色
- **进度条 + 日志** - 后台线程处理，实时进度更新

## 产品架构

```
excel_page_break_tool/
├── main.py                      # 程序入口
├── config.py                    # 全局配置（颜色、命名、UI参数）
├── requirements.txt             # 依赖
├── core/
│   ├── excel_processor.py       # 核心处理（分页符+背景色+备份）
│   ├── serial_number_detector.py # 序号检测器（2/3/4位自动识别）
│   └── worker.py                # 后台工作线程
├── ui/
│   ├── main_window.py           # 主窗口（信号槽连接）
│   ├── left_panel.py            # 左侧控制面板
│   └── preview_panel.py         # 右侧预览画布
├── tests/
│   └── test_serial_detector.py  # 单元测试（6项）
├── docs/
│   ├── PRD_V1.md                # 产品需求 V1
│   ├── PRD_V2.md                # 产品需求 V2
│   ├── DEVELOPMENT_GUIDE.md     # 开发指南
│   └── legacy_process_excel.py  # V1 命令行版本（归档）
└── test_flow.py                 # 集成测试（7项）
```

## 技术栈

| 层级 | 技术 | 版本 |
|------|------|------|
| 开发语言 | Python | 3.10+ |
| GUI框架 | PyQt6 | 6.4+ |
| Excel处理 | openpyxl | 3.1+ |
| 打包 | PyInstaller | 6.0+ |

## 快速开始

安装依赖：

    pip install -r requirements.txt

启动应用：

    python main.py

运行测试：

    # 单元测试
    python -m unittest tests.test_serial_detector -v

    # 集成测试
    python test_flow.py

## 数据处理流程

输入文件格式：6列 Excel（顺序号、责任者、题名、日期、页号、备注）

分页规则：当顺序号从非1变为1时，在该行上方插入分页符

输出：
- 文件名：{原文件名}_{YYYYMMDD}_V1.0.xlsx
- 备份：{原文件名}_原始备份.xlsx

## 许可证

MIT