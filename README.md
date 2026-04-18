# 📄 Excel 分页符工具 (Excel Page Break Tool)

> 档案管理专用工具 — 自动识别卷内目录 Excel 中的案卷边界，批量插入分页符，一键下载成品文件。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Pure Frontend](https://img.shields.io/badge/Pure_Frontend-✅-blue)]()
[![Zero Install](https://img.shields.io/badge/Zero_Install-打开即用-green)]()

---

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| 📤 拖拽上传 | 支持 .xlsx / .xls 文件拖拽或点击上传 |
| 🔍 智能识别 | 自动检测顺序号=1的行，精准定位分页点 |
| 📊 统计概览 | 实时展示卷数、文件总数、分页点数量 |
| 📑 分卷浏览 | 下拉选择 / 上下翻页，快速定位到任意卷 |
| 📥 一键导出 | 下载带分页符的 Excel 文件，直接打印 |
| 🖨️ 打印预览 | 模拟分页打印效果，所见即所得 |
| 🤖 AI 助手 | 预留 DeepSeek API 接口，支持智能问答 |

---

## 🚀 快速使用

### 方式一：直接打开（推荐）

```bash
# 浏览器直接打开 index.html，无需任何安装
open index.html
```

### 方式二：本地服务器

```bash
# Python
python3 -m http.server 8080

# Node.js
npx serve .
```

### 方式三：后端模式（可选）

详见 [backend/README.md](backend/README.md)

---

## 📋 使用流程

```
第一步 → 上传 Excel 文件（拖拽或点击）
         ↓
第二步 → 系统自动分析，展示统计和预览
         ↓
第三步 → 点击下载，获取带分页符的成品
```

---

## 🏗️ 系统架构

```
┌─────────────────────────────────────────────────┐
│                  index.html                      │
│              （单文件纯前端应用）                  │
├─────────────┬──────────────┬────────────────────┤
│  上传模块    │   解析引擎    │    导出模块         │
│  Upload     │   Parser     │    Export           │
│  (Drag&Drop)│   (SheetJS)  │    (SheetJS)        │
├─────────────┼──────────────┼────────────────────┤
│  分页检测    │   卷管理器    │    AI 助手          │
│  BreakFinder│  VolManager  │   (DeepSeek 预留)   │
├─────────────┴──────────────┴────────────────────┤
│                  浏览器本地运行                    │
│            无需服务器 · 无需安装依赖               │
└─────────────────────────────────────────────────┘
         │                    │
         │  可选               │  可选
         ▼                    ▼
┌─────────────────┐  ┌─────────────────┐
│  backend/        │  │  DeepSeek API   │
│  Flask + openpyxl│  │  (AI 智能分析)   │
│  (Python 后端)   │  │                 │
└─────────────────┘  └─────────────────┘
```

### 核心算法

**分页检测规则**：当顺序号从 `≠1` 变为 `=1` 时，在该行上方插入分页符。

```javascript
// 简化的分页检测逻辑
for (let i = 1; i < data.length; i++) {
    const curr = parseInt(data[i][0]);  // 当前行顺序号
    const prev = parseInt(data[i-1][0]); // 上一行顺序号
    if (curr === 1 && prev !== 1) {
        pageBreaks.push(i);  // 此处插入分页符
    }
}
```

---

## 📁 项目结构

```
excel-page-break-tool/
├── index.html              # 主应用（纯前端，单文件）
├── backend/
│   ├── processor.py        # Python 后端（Flask + openpyxl）
│   └── requirements.txt    # Python 依赖
├── docs/
│   └── PRD.md              # 产品需求文档
├── .gitignore
├── LICENSE                 # MIT
└── README.md               # 本文件
```

---

## 🔧 技术栈

| 组件 | 技术 | 说明 |
|------|------|------|
| **前端** | HTML + CSS + Vanilla JS | 单文件，零依赖 |
| **Excel 解析** | [SheetJS](https://sheetjs.com/) (CDN) | 浏览器端读写 Excel |
| **后端（可选）** | Flask + openpyxl | Python 服务端处理 |
| **AI（预留）** | DeepSeek API | 智能分析问答 |

---

## 📊 数据格式要求

Excel 文件需满足以下结构：

| 列 | 说明 | 示例 |
|----|------|------|
| 顺序号 | 卷内文件序号，每卷从 1 开始 | 1, 2, 3, ..., 1, 2, ... |
| 责任者 | 文件责任主体 | XX公司 |
| 题名 | 文件标题 | XX产权交易鉴证书 |
| 日期 | 文件日期 | 20240103 |
| 页号 | 所在页码 | 1 |
| 备注 | 备注信息（可空） | - |

**关键**：前 2 行为表头，第 3 行开始为数据。顺序号列必须是第一列。

---

## 🤝 贡献

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/xxx`)
3. 提交修改 (`git commit -m 'Add xxx'`)
4. 推送分支 (`git push origin feature/xxx`)
5. 提交 Pull Request

---

## 📝 License

[MIT](LICENSE) © 2026 luomeng119
