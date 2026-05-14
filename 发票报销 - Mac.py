import sys
import os
import re
import fitz
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QListWidgetItem, QFileDialog,
    QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent


# ================== PDF 提取逻辑 ==================
def extract_invoice_data(path):
    """从单张 + 循环"""
    try:
        doc = fitz.open(path)
        full_text = ""
        for page in doc:
            full_text += page.get_text()
        doc.close()
    except Exception as e:
        print(f"[错误] 无法读取 {path}: {e}")
        return None

    data = {"文件名": os.path.basename(path)}


    # ==================== 付款人 ====================
    filename = os.path.basename(path)  
    name_without_ext = os.path.splitext(filename)[0] 

    parts = name_without_ext.split('-')
    candidate = parts[-1]  
    data["付款人"] = candidate if candidate else ""


    # ======== 发票号码：20位数字 ========
    m = re.search(r"(\d{20})", full_text)
    data["发票号码"] = m.group(1) if m else "⛔️"


    # ======== 开票日期 ========
    # 搜索年月日
    m = re.search(r"(\d{4}年\d{2}月\d{2}日)", full_text)
    data["开票日期"] = m.group(1) if m else "⛔️"


    # ======== 供应商 ========
    # 寻找特定规律，供应商在发票号前面
    """ ======== 优化空间 ======== """
    # 搜索方式

    try:
        pairs = re.findall(r"([^\n]{2,30}?)\s*\n\s*([A-Z0-9]{15,20})", full_text)

        valid_pairs = [(name, tax) for name, tax in pairs 
                if re.search(r'[A-Z]', tax) and re.search(r'\d', tax) or re.search(r'\d{18}', tax) ]

        data["供应商"] = valid_pairs[2][0] if len(valid_pairs) >= 3 else valid_pairs[1][0]
    except Exception as e:
        print(f'   ⛔️ {path} 无供应商')
        data["供应商"] = "⛔️"


    # ======== 数量 + 单位 ======== #
    """ ========= 注释 ========== """
    # 代码识别不同发票的差异大 - 数量一般集中在单位后四行内
    # 先搜索单位后，遍历单位后四行的数值，只搜寻整数
    # 输出整数的同时输出相应的单位
    """ ======== 优化空间 ======== """
    # 单位的识别
    # 搜索方式

    lines = [t.strip() for t in re.split(r'\s+', full_text) if t.strip()]
    units = ["条", "套", "个", "盒", "卷", "支", "包", "批", "台", "件",
             "根", "只", "把", "米", "次", "份", "粒", "克", "斤", "张"]

    quantity = 0
    for i, line in enumerate(lines):
        if line in units:
            next_line_i = i+1
            for i in range (next_line_i,next_line_i+4):
                next_line = lines[i]
                if re.match(r"^\d+$", next_line):  # 纯整数
                    quantity += int(next_line)
                    data["数量"] = quantity
                    data["单位"] = line if line not in r"套包盒批份次" else line + "❌"
 

    # ======== 金额：提取所有 ¥开头的金额 ========
    #查找以¥开头的数字，取最大值

    amounts = re.findall(r"¥\s*(\d+\.\d{2})", full_text)
    data["金额"] = max(amounts) if amounts else ""


    # ======== 单价 ======== 
    # 金额 / 数量

    try:
        ave = round(float(max(amounts)) / int(quantity),2)
    except Exception as e:
        print(f'   ⛔️ {path} 无单价')
        ave = 0
    data["单价"] = ave


    # ======== 类型：易耗品 + 低值判断 ======== 
    # 通过单价判断
    data["类型"] = "易耗品" if ave < 200 else "❗️低值"


    # ======== 项目名称 ======== 
    # 文件名用 “-” 分割，取倒一

    try:
        m = re.findall(r"\*([\u4e00-\u9fa5A-Z0-9]+)\*", full_text)
        data["项目名称"] = m[0]
    except Exception as e:
        print(f'   ⛔️ {path} 无项目名称')
        data["项目名称"] = "⛔️"


    # ======== 规格型号 ========
    # 文件名用 “-” 分割，取倒一

    try:
        a = re.findall(r"\*[\u4e00-\u9fa5A-Z0-9]+\*[\w\u4e00-\u9fa5A-Z0-9]+\n([\w\u4e00-\u9fa5A-Z0-9]+)", full_text)
        candidate = ""
        for can in a:
            candidate += "\n" + can
        if len(candidate) > 4:
            data["规格型号"] = candidate
        else:
            data["规格型号"] = ""
        # ["1","2","3","4","5","6","7","8","9","0"]
    except Exception as e:
        print(f'   ⛔️ {path} 规格型号')
        data["项目名称"] = "⛔️"


    return data



# ========== batch_invoices_to_excel ========== #
#遍历文件 + 提取信息

def process_pdfs_to_excel(file_paths, save_path):

    records = []
    for path in file_paths:
        record = extract_invoice_data(path)
        if record:
            records.append(record)

    if not records:
        return None

    # 定义列顺序：同一种data分到一列
    columns = [
        "文件名", "类型", "项目名称", "单价", "数量", 
        "单位", "金额", "发票号码", 
        "开票日期", "供应商", "规格型号","付款人"
    ]

    df = pd.DataFrame(records)
    # 确保列顺序统一，缺失列补空字符串
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    df = df[columns]

    df["金额"] = pd.to_numeric(df["金额"], errors="coerce")

    summary = df.groupby("付款人", as_index=False)["金额"].sum()

    summary["金额合计"] = summary["金额"].apply(lambda x: f"¥{x:.2f}")

    summary = summary[["付款人", "金额合计"]]

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
    
        # Sheet 1：原始明细
        df.to_excel(writer, sheet_name="发票明细", index=False)
        
        # Sheet 2：每人合计
        summary.to_excel(writer, sheet_name="按人汇总", index=False)

    # 生成Excel
    return save_path


# ================== 后台线程 ==================
class Worker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, file_paths, save_path):
        super().__init__()
        self.file_paths = file_paths
        self.save_path = save_path

    def run(self):
        try:
            total = len(self.file_paths)
            for i, _ in enumerate(self.file_paths, 1):
                self.progress.emit(int(i / total * 100))

            result = process_pdfs_to_excel(self.file_paths, self.save_path)
            if result:
                self.finished.emit(result)
            else:
                self.error.emit("未能从PDF中提取到有效数据")
        except Exception as e:
            self.error.emit(str(e))


# ================== 主窗口 ==================
class InvoiceWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("机器人实验室发票报销小程序 · 3.4")
        self.setGeometry(543, 124, 600, 800)
        self.setAcceptDrops(True)

        self.pdf_files = []
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 20, 20, 20)

        # 标题
        title = QLabel("低值及耗材类采购申请明细")
        title_font = QFont("PingFang SC", 20, QFont.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #B22222; margin-bottom: 10px;")
        layout.addWidget(title)

        # 拖拽区
        self.drop_label = QLabel("拖拽 PDF 发票文件\n或点击下方按钮选择文件")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setMinimumHeight(100)
        self.drop_label.setStyleSheet("""
            QLabel {
                background-color: #fdf2f2;
                border: 2px solid #B22222;
                border-radius: 11px;
                color: #666666;
                font-size: 14px;
                padding: 20px;
            }
        """)
        layout.addWidget(self.drop_label)

        # 按钮行
        btn_layout = QHBoxLayout()

        self.btn_add = QPushButton("✚ 选择 PDF 文件")
        self.btn_add.setStyleSheet(self._btn_style("#B22222"))
        self.btn_add.clicked.connect(self.select_files)

        self.btn_clear = QPushButton("🗑️ 清空列表")
        self.btn_clear.setStyleSheet(self._btn_style("#666666"))
        self.btn_clear.clicked.connect(self.clear_files)

        btn_layout.addWidget(self.btn_add)
        btn_layout.addWidget(self.btn_clear)
        layout.addLayout(btn_layout)

        # 文件列表
        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("""
            QListWidget {
                border: 1px solid #dddddd;
                border-radius: 6px;
                padding: 5px;
                font-size: 13px;
            }
            QListWidget::item {
                padding: 6px;
                border-bottom: 1px solid #f0f0f0;
            }
        """)
        layout.addWidget(self.list_widget)

        # 进度条
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #dddddd;
                border-radius: 4px;
                text-align: center;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #B22222;
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.progress)

        # 生成按钮
        self.btn_run = QPushButton("生成 Excel 明细")
        self.btn_run.setStyleSheet(self._btn_style("#B22222", height=44))
        run_font = QFont("PingFang SC", 12, QFont.Bold)
        self.btn_run.setFont(run_font)
        self.btn_run.clicked.connect(self.run_extraction)
        layout.addWidget(self.btn_run)

        # 状态栏
        self.status = QLabel("KAISEN ©")
        self.status.setAlignment(Qt.AlignCenter)
        self.status.setStyleSheet("color: #999999; font-size: 12px; margin-top: 8px;")
        layout.addWidget(self.status)

    def _btn_style(self, color, height=36):
        return f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 15px;
                height: {height}px;
            }}
            QPushButton:hover {{
                background-color: #8B0000;
            }}
            QPushButton:pressed {{
                background-color: #660000;
            }}
            QPushButton:disabled {{
                background-color: #cccccc;
            }}
        """

    # ================== 拖拽 ==================
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith('.pdf') and path not in self.pdf_files:
                self.pdf_files.append(path)
                item = QListWidgetItem(f"📄 {os.path.basename(path)}")
                item.setToolTip(path)
                self.list_widget.addItem(item)
        self._update_status()

    # ================== 文件操作 ==================
    def select_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "选择 PDF 发票文件", "", "PDF Files (*.pdf)"
        )
        for path in paths:
            if path not in self.pdf_files:
                self.pdf_files.append(path)
                item = QListWidgetItem(f"📄 {os.path.basename(path)}")
                item.setToolTip(path)
                self.list_widget.addItem(item)
        self._update_status()

    def clear_files(self):
        self.pdf_files.clear()
        self.list_widget.clear()
        self._update_status()

    def _update_status(self):
        count = len(self.pdf_files)
        self.status.setText(f"已加载 {count} 个 PDF 文件" if count else "KAISEN ©")
        self.drop_label.setText(
            f"已拖拽 {count} 个文件\n可继续拖拽追加" if count else "拖拽 PDF 文件到此处\n或点击下方按钮选择文件"
        )

    # ================== 核心处理 ==================
    def run_extraction(self):
        if not self.pdf_files:
            QMessageBox.warning(self, "提示", "请先拖拽或选择 PDF 文件")
            return

        default_name = f"{datetime.now().strftime('%Y%m')}低值及耗材类采购申请明细.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存 Excel 报表", default_name, "Excel Files (*.xlsx)"
        )
        if not save_path:
            return

        self.btn_run.setEnabled(False)
        self.btn_add.setEnabled(False)
        self.btn_clear.setEnabled(False)
        self.status.setText("正在提取数据，请稍候...")
        self.progress.setValue(0)

        self.worker = Worker(self.pdf_files, save_path)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_success)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_success(self, path):
        self.btn_run.setEnabled(True)
        self.btn_add.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.status.setText(f"✅ 完成：{os.path.basename(path)}")

        reply = QMessageBox.question(
            self, "处理完成", f"Excel 已保存到：\n{path}\n\n是否立即打开所在文件夹？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes
        )
        if reply == QMessageBox.Yes:
            # macOS 用 open 命令
            os.system(f'open "{os.path.dirname(path)}"')

    def on_error(self, msg):
        self.btn_run.setEnabled(True)
        self.btn_add.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.progress.setValue(0)
        self.status.setText("❌ 处理失败")
        QMessageBox.critical(self, "错误", msg)


# ================== MAIN ==================
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # 修复：使用 macOS 字体
    font = QFont("PingFang SC", 10)
    if not QFont(font).exactMatch():
        font = QFont("Helvetica Neue", 10)
    app.setFont(font)

    window = InvoiceWindow()
    window.show()
    sys.exit(app.exec_())