# 发票报销助手 

<p align="center">
  <img src="https://img.shields.io/badge/version-4.0-red" alt="Version">
  <img src="https://img.shields.io/badge/python-3.13%2B-blue" alt="Python">
  <img src="https://img.shields.io/badge/platform-macOS%20%7C%20Windows-lightgrey" alt="Platform">
  <img src="https://img.shields.io/badge/license-MIT-green" alt="License">
</p>

<p align="center">
  <b>一键提取 PDF 发票信息，自动生成 Excel 报销报表</b>
</p>

---

## 📋 项目简介

本项目是 **SJTU VEX 机器人实验室** 内部使用的发票报销辅助工具，基于 Python + PyQt5 开发，支持 **macOS** 和 **Windows** 双平台。它可以批量读取 PDF 格式的电子发票，自动提取关键字段并生成结构化的 Excel 报销明细表，大幅提升财务报销效率。

> 💡 从 Jupyter Notebook 原型开始，历经 4 个版本迭代，现已打包为可直接运行的桌面应用程序。

---

## ✨ 核心功能

| 功能 | 说明 |
|:---|:---|
| 📄 **批量 PDF 导入** | 支持文件选择对话框或拖拽方式批量导入发票 |
| 🔍 **智能信息提取** | 自动识别发票号码、开票日期、供应商、金额、数量、单位等 |
| 🏷️ **自动类型判断** | 根据单价智能区分「易耗品」与「低值」资产 |
| 👤 **付款人识别** | 通过文件名后缀自动识别付款人（如 `发票-张三.pdf`） |
| 📊 **Excel 导出** | 自动生成含「发票明细」和「按人汇总」双 Sheet 的报表 |
| 🎛️ **自定义列顺序** | Mac 版支持拖拽调整输出列顺序、删除不需要的列 |
| 🎨 **现代化 UI** | 简洁美观的图形界面，支持进度条实时反馈 |

### 提取字段清单

- 文件名
- 类型（易耗品 / 低值）
- 项目名称
- 单价
- 数量
- 单位
- 金额
- 发票号码（20 位）
- 开票日期
- 供应商
- 规格型号
- 付款人

---

## 🖼️ 界面预览

```
┌─────────────────────────────────────────┐
│     低值及耗材类采购申请明细                │
│     机器人实验室 发票报销小程序v4.0         │
├─────────────────────────────────────────┤
│                                         │
│   ┌─────────────────────────────────┐   │
│   │  拖拽 PDF 发票文件到此处           │   │
│   │  或点击下方按钮选择文件             │   │
│   └─────────────────────────────────┘   │
│                                         │
│   [  选择 PDF 文件  ]  [  清空列表  ]      │
│                                         │
│   待处理文件                    0 个文件   │
│   ┌─────────────────────────────────┐   │
│   │                                 │   │
│   │      （文件列表区域）              │   │
│   │                                 │   │
│   └─────────────────────────────────┘   │
│                                         │
│   ████████░░░░░░░░░░  进度条             │
│                                         │
│   [      生成 Excel 明细      ]          │
│                                         │
└─────────────────────────────────────────┘
```

---

## 🚀 快速开始

### 方式一：直接运行 Python 脚本

#### 1. 克隆仓库

```bash
git clone https://github.com/kaisch314-ops/invoice.git
```

#### 2. 安装依赖

```bash
pip install PyQt5 PyMuPDF pandas openpyxl
```

#### 3. 运行程序

**macOS:**
```bash
python "发票报销 - Mac.py"
```

**Windows:**
```bash
python "发票报销 - Win.py"
```

### 方式二：运行打包好的应用（macOS）

双击 `发票报销4.0.app` 即可启动，无需安装 Python 环境。

> ⚠️ 如果提示「无法打开」，请前往 **系统设置 → 隐私与安全性** 中允许运行。

---

## 📁 项目结构

```
invoice-reimbursement/
├── 发票报销 - Mac.py          # macOS 版主程序（v4.0，支持列自定义）
├── 发票报销 - Win.py          # Windows 版主程序（v3.4）
├── 发票报销4.0.app/           # macOS 打包应用（PyInstaller）
├── SJTU-VEX发票报销.ipynb      # Jupyter Notebook 原型与开发记录
├── 202605低值及耗材类采购申请明细.xlsx   # 输出示例
└── README.md
```

---

## 🛠️ 打包为独立应用

### macOS

```bash
# 安装 pyinstaller
pip install pyinstaller

# 打包为 .app
pyinstaller --windowed --onefile --name "发票报销4.0" \
  --icon icon.icns \
  "发票报销 - Mac.py"
```

### Windows

```bash
pyinstaller --windowed --onefile --name "发票报销" \
  --icon icon.ico \
  "发票报销 - Win.py"
```

---

## 📝 使用说明

1. **准备发票**
   - 将所有 PDF 发票放入同一文件夹
   - **文件名格式建议**：`发票内容-付款人姓名.pdf`（如 `螺丝刀套装-张三.pdf`）

2. **导入文件**
   - 点击「选择 PDF 文件」按钮，或直接将文件拖拽到窗口

3. **确认列顺序**（Mac 版）
   - 在弹出的对话框中拖拽调整列顺序
   - 点击 ✕ 删除不需要的列

4. **生成报表**
   - 点击「生成 Excel 明细」
   - 选择保存位置，默认文件名为 `YYYYMM低值及耗材类采购申请明细.xlsx`

5. **查看结果**
   - **Sheet 1「发票明细」**：所有发票的详细字段
   - **Sheet 2「按人汇总」**：按付款人汇总的金额合计

---

## ⚙️ 技术栈

- **GUI 框架**: [PyQt5](https://www.riverbankcomputing.com/software/pyqt/)
- **PDF 解析**: [PyMuPDF (fitz)](https://github.com/pymupdf/PyMuPDF)
- **数据处理**: [pandas](https://pandas.pydata.org/)
- **Excel 导出**: [openpyxl](https://openpyxl.readthedocs.io/)
- **打包工具**: [PyInstaller](https://pyinstaller.org/)

---

## 版本演进

| 版本 | 形式 | 主要特性 |
|:---:|:---|:---|
| 1.0 | Jupyter Notebook | 基础 PDF 提取 + Excel 输出 |
| 2.0 | Python CLI | 批量处理、路径自定义 |
| 3.0 | PyQt5 GUI | 图形界面、拖拽导入、双平台支持 |
| 3.4 | PyQt5 GUI | 优化文件大小 |
| **4.0** | **PyQt5 GUI + .app** | **列顺序自定义、UI 焕新、windows + macOS 打包** |

---

## 贡献指南

欢迎提交 Issue 或 Pull Request！

- 发现发票解析错误？请提供 **脱敏后的 PDF 样本**
- 有新功能建议？请在 Issue 中详细描述使用场景
- 想适配 Linux？欢迎 fork 并提交 PR

---

## 📄 许可证

本项目采用 [MIT License](LICENSE) 开源协议。

---

<p align="center">
  <sub>Made by KAISEN · SJTU VEX Robotics Lab</sub>
</p>
